#!/usr/bin/env python3
"""
POS Data Cleaner 2 - Web API Backend
Integrates the production-ready processing engine (98.9% accuracy) with the web interface
"""

import os
import sys
import json
import uuid
import time
import shutil
import subprocess
import io
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Any, Optional

from flask import Flask, request, jsonify, send_file, make_response
from flask_cors import CORS
from werkzeug.utils import secure_filename
import pandas as pd
from dotenv import load_dotenv
from windows_utils import safe_print, set_console_encoding

# Set console encoding for Windows compatibility
set_console_encoding()

# Load environment variables from parent directory
from pathlib import Path
env_path = Path(__file__).parent.parent / '.env'
load_dotenv(dotenv_path=env_path)

# Import storage service
from storage_service import storage_service
# from file_scanner import FileScanner  # COMMENTED OUT: Module doesn't exist, endpoint not used by frontend

# Helper function to safely serialize pandas data to JSON
def safe_serialize_dict(data_dict):
    """Convert pandas/numpy data types to JSON-serializable Python types"""
    result = {}
    for key, value in data_dict.items():
        if pd.isna(value):
            result[key] = None
        elif hasattr(value, 'dtype'):
            # Handle numpy/pandas numeric types
            if pd.api.types.is_integer_dtype(value):
                result[key] = int(value)
            elif pd.api.types.is_float_dtype(value):
                result[key] = float(value)
            else:
                result[key] = str(value)
        else:
            result[key] = value
    return result

# Migrate Excel config files to CSV on startup
def migrate_config_to_csv():
    """Migrate config files from Excel to CSV format"""
    try:
        # Check if Excel file exists but CSV doesn't
        if storage_service.file_exists('config/aspects.xlsx'):
            aspects_data = storage_service.read_excel('config/aspects.xlsx')
            if aspects_data is not None and not storage_service.file_exists('config/aspects.csv'):
                safe_print("🔄 Migrating aspects from Excel to CSV...")
                success = storage_service.write_csv(aspects_data, 'config/aspects.csv')
                if success:
                    safe_print("✅ Successfully migrated aspects to CSV")
                else:
                    safe_print("❌ Failed to migrate aspects to CSV")
            else:
                safe_print("📋 Aspects CSV already exists or no Excel data to migrate")
    except Exception as e:
        safe_print(f"⚠️ Error during config migration: {e}")

# Run migration on startup
migrate_config_to_csv()

def generate_unique_id():
    """Generate a unique ID for database records"""
    return int(time.time() * 1000000) % 2147483647  # Generate int ID within PostgreSQL int range

def generate_checklist_id(year, row_number):
    """Generate checklist ID in format: last two digits of year + row number
    Example: 2024 row 1 -> 241, 2025 row 10 -> 2510"""
    year_digits = year % 100  # Get last two digits of year
    return int(f"{year_digits}{row_number}")

# Project root for subprocess calls to the working core system
project_root = str(Path(__file__).parent.parent.parent)

app = Flask(__name__)
# Enable CORS for React frontend with exposed headers (allow all localhost ports for development)
CORS(app,
     origins=["*"],  # Allow all origins for development
     methods=["GET", "POST", "PUT", "DELETE", "OPTIONS"],
     allow_headers=["Content-Type", "Authorization", "X-Requested-With"],
     expose_headers=['Content-Disposition', 'Content-Type', 'Content-Length'],
     supports_credentials=True)

# Configuration
UPLOAD_FOLDER = Path(__file__).parent / 'uploads'
OUTPUT_FOLDER = Path(__file__).parent / 'outputs'
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'pdf', 'png', 'jpg', 'jpeg', 'txt', 'md', 'markdown'}

# Ensure directories exist
UPLOAD_FOLDER.mkdir(exist_ok=True)
OUTPUT_FOLDER.mkdir(exist_ok=True)

app.config['UPLOAD_FOLDER'] = str(UPLOAD_FOLDER)
app.config['OUTPUT_FOLDER'] = str(OUTPUT_FOLDER)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500MB max file size (increased from 50MB)

def allowed_file(filename: str) -> bool:
    """Check if file extension is allowed."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_file_type(filename: str) -> str:
    """Determine file type from extension."""
    ext = filename.rsplit('.', 1)[1].lower()
    if ext in {'xlsx', 'xls'}:
        return 'excel'
    elif ext == 'pdf':
        return 'pdf'
    elif ext in {'png', 'jpg', 'jpeg'}:
        return 'image'
    return 'unknown'

# Global error handlers
@app.errorhandler(413)
def request_entity_too_large(error):
    """Handle file too large errors"""
    safe_print(f"❌ ERROR 413: Request Entity Too Large")
    return jsonify({
        'error': 'File terlalu besar. Maksimal 50MB per file.'
    }), 413

@app.errorhandler(Exception)
def handle_exception(error):
    """Handle all unhandled exceptions"""
    safe_print(f"❌ UNHANDLED EXCEPTION: {type(error).__name__}: {error}")
    import traceback
    safe_print(f"❌ TRACEBACK: {traceback.format_exc()}")
    return jsonify({
        'error': f'Internal server error: {str(error)}'
    }), 500

@app.route('/api/health', methods=['GET'])
def health_check():
    """Health check endpoint."""
    return jsonify({
        'status': 'healthy',
        'service': 'POS Data Cleaner 2 API',
        'version': '2.0.0',
        'timestamp': datetime.now().isoformat()
    })

# COMMENTED OUT: FileScanner module doesn't exist, endpoint not used by frontend
# @app.route('/api/scan-and-sync-files', methods=['POST'])
# def scan_and_sync_files():
#     """
#     Scan file storage and synchronize with database
#
#     For Data Engineers: Drop files into storage, then hit this endpoint to sync
#
#     Expected file structure:
#     data/gcg-documents/{year}/{subdirektorat}/{checklist_id}/{filename}
#
#     Returns:
#         JSON with scan statistics
#     """
#     try:
#         # Get optional parameters
#         data = request.get_json() or {}
#         uploaded_by = data.get('uploaded_by', 'File Scanner - API')
#
#         # Initialize scanner
#         scanner = FileScanner()
#
#         # Perform scan and sync
#         results = scanner.scan_and_sync(uploaded_by=uploaded_by)
#
#         return jsonify({
#             'success': True,
#             'message': 'File storage scanned and synchronized successfully',
#             'results': results
#         }), 200
#
#     except Exception as e:
#         safe_print(f"Error scanning and syncing files: {e}")
#         import traceback
#         traceback.print_exc()
#         return jsonify({
#             'success': False,
#             'error': f'Failed to scan and sync files: {str(e)}'
#         }), 500

@app.route('/api/upload', methods=['POST'])
def upload_file():
    """
    Upload and process GCG assessment document.
    
    Expected form data:
    - file: The document file
    - checklistId: (optional) Associated checklist item ID
    - year: (optional) Assessment year
    - aspect: (optional) GCG aspect
    """
    try:
        safe_print(f"🔧 DEBUG: Upload request received")
        safe_print(f"🔧 DEBUG: Request files: {list(request.files.keys())}")
        
        # Check if file is present
        if 'file' not in request.files:
            safe_print(f"🔧 DEBUG: No file in request")
            return jsonify({'error': 'No file provided'}), 400
        
        file = request.files['file']
        safe_print(f"🔧 DEBUG: File received: {file.filename}")
        
        if file.filename == '':
            safe_print(f"🔧 DEBUG: Empty filename")
            return jsonify({'error': 'No file selected'}), 400
        
        if not allowed_file(file.filename):
            safe_print(f"🔧 DEBUG: File type not allowed: {file.filename}")
            return jsonify({'error': 'File type not allowed'}), 400
        
        safe_print(f"🔧 DEBUG: File validation passed")
        
        try:
            safe_print(f"🔧 DEBUG: Starting file processing...")
            # Generate unique filename
            file_id = str(uuid.uuid4())
            safe_print(f"🔧 DEBUG: Generated file_id: {file_id}")
            original_filename = secure_filename(file.filename)
            filename_parts = original_filename.rsplit('.', 1)
            unique_filename = f"{file_id}_{filename_parts[0]}.{filename_parts[1]}"
            
            # Save uploaded file
            input_path = UPLOAD_FOLDER / unique_filename
            file.save(str(input_path))
            
            # Generate output filename
            output_filename = f"processed_{file_id}_{filename_parts[0]}.xlsx"
            output_path = OUTPUT_FOLDER / output_filename
            
            # Get metadata from form
            checklist_id = request.form.get('checklistId')
            year = request.form.get('year')
            aspect = request.form.get('aspect')
            
            # Process the document using production system
            file_type = get_file_type(original_filename)
            
            if file_type == 'excel':
                safe_print(f"🔧 DEBUG: Processing Excel file using core system (subprocess)...")
                processing_result = None  # Force use of subprocess method
                
                # DISABLED: Accurate processing has pandas column selection issue
                # Fall through to subprocess method which works perfectly
            
            # Use subprocess method for all file types (Excel, PDF, Image)
            if file_type in ['excel', 'pdf', 'image']:
                safe_print(f"🔧 DEBUG: Processing {file_type} file using core system...")
                
                try:
                    import time
                    start_time = time.time()
                    
                    # Call the working core system directly as subprocess
                    cmd = [
                        sys.executable, "main_new.py",
                        "-i", str(input_path),
                        "-o", str(output_path),
                        "-v"
                    ]
                    
                    safe_print(f"🔧 DEBUG: Running command: {' '.join(cmd)}")
                    safe_print(f"🔧 DEBUG: Working directory: {project_root}")
                    
                    result = subprocess.run(
                        cmd,
                        cwd=project_root,
                        capture_output=True,
                        text=True,
                        timeout=180  # 3 minute timeout for OCR processing
                    )
                    
                    end_time = time.time()
                    safe_print(f"🔧 DEBUG: Core system completed in {end_time - start_time:.2f} seconds")
                    safe_print(f"🔧 DEBUG: Return code: {result.returncode}")
                    safe_print(f"🔧 DEBUG: STDOUT: {result.stdout}")
                    if result.stderr:
                        safe_print(f"🔧 DEBUG: STDERR: {result.stderr}")
                    
                    if result.returncode == 0:
                        processing_result = {
                            'success': True,
                            'method': f'{file_type}_processing',
                            'message': 'Processing completed successfully',
                            'stdout': result.stdout,
                            'processing_time': f"{end_time - start_time:.2f}s"
                        }
                    else:
                        processing_result = {
                            'success': False,
                            'method': f'{file_type}_processing',
                            'error': f'Core system failed with code {result.returncode}',
                            'stdout': result.stdout,
                            'stderr': result.stderr
                        }
                    
                except subprocess.TimeoutExpired:
                    processing_result = {
                        'success': False,
                        'method': f'{file_type}_processing',
                        'error': 'Processing timeout (3 minutes exceeded)'
                    }
                except Exception as e:
                    safe_print(f"🔧 DEBUG: EXCEPTION in subprocess call: {e}")
                    import traceback
                    safe_print(f"🔧 DEBUG: Full traceback: {traceback.format_exc()}")
                    processing_result = {
                        'success': False,
                        'method': f'{file_type}_processing',
                        'error': f'Subprocess failed: {str(e)}'
                    }
            
            else:
                processing_result = {
                    'success': False,
                    'error': f'Unsupported file type: {file_type}',
                    'method': 'unsupported'
                }
        
        except Exception as proc_error:
            processing_result = {
                'success': False,
                'error': f'Processing failed: {str(proc_error)}',
                'method': 'processing_error'
            }
        
        # Load processed results if successful
        extracted_data = None
        if processing_result['success'] and output_path.exists():
            try:
                # Read the processed Excel file
                df = pd.read_excel(str(output_path))
                safe_print(f"🔧 DEBUG: Loaded DataFrame with {len(df)} rows")
                safe_print(f"🔧 DEBUG: DataFrame columns: {list(df.columns)}")
                safe_print(f"🔧 DEBUG: DataFrame head:\n{df.head()}")
                
                # Extract key metrics
                indicator_rows = df[df['Type'] == 'indicator'] if 'Type' in df.columns else df
                subtotal_rows = df[df['Type'] == 'subtotal'] if 'Type' in df.columns else pd.DataFrame()
                total_rows = df[df['Type'] == 'total'] if 'Type' in df.columns else pd.DataFrame()
                safe_print(f"🔧 DEBUG: Found {len(indicator_rows)} indicator rows")
                
                extracted_data = {
                    'total_rows': int(len(df)),
                    'indicators': int(len(indicator_rows)),
                    'subtotals': int(len(subtotal_rows)),
                    'totals': int(len(total_rows)),
                    'year': str(df['Tahun'].iloc[0]) if len(df) > 0 and pd.notna(df['Tahun'].iloc[0]) else None,
                    'penilai': str(df['Penilai'].iloc[0]) if len(df) > 0 and pd.notna(df['Penilai'].iloc[0]) else None,
                    'format_type': 'DETAILED' if len(df) > 20 else 'BRIEF',
                    'processing_status': 'success'
                }
                
                # Extract ALL indicator data (not just samples)
                if len(indicator_rows) > 0:
                    all_indicators = []
                    for _, row in indicator_rows.iterrows():
                        all_indicators.append({
                            'no': int(row['No']) if pd.notna(row['No']) else 0,
                            'section': str(row['Section']) if pd.notna(row['Section']) else '',
                            'description': str(row['Deskripsi']) if pd.notna(row['Deskripsi']) else '',
                            'jumlah_parameter': int(row['Jumlah_Parameter']) if pd.notna(row['Jumlah_Parameter']) else 0,
                            'bobot': float(row['Bobot']) if pd.notna(row['Bobot']) else 100.0,
                            'skor': float(row['Skor']) if pd.notna(row['Skor']) else 0.0,
                            'capaian': float(row['Capaian']) if pd.notna(row['Capaian']) else 0.0,
                            'penjelasan': str(row['Penjelasan']) if pd.notna(row['Penjelasan']) else 'Sangat Kurang'
                        })
                    extracted_data['sample_indicators'] = all_indicators
                    
                # Add sheet analysis for XLSX files and extract BRIEF data for aspect summary
                if file_type == 'excel':
                    try:
                        # Read Excel file to analyze sheets
                        excel_file = pd.ExcelFile(str(input_path))
                        sheet_names = excel_file.sheet_names
                        
                        sheet_analysis = {
                            'total_sheets': len(sheet_names),
                            'sheet_names': sheet_names,
                            'sheet_types': {}
                        }
                        
                        brief_sheet_data = None
                        
                        # Analyze each sheet to determine if it's BRIEF or DETAILED
                        for sheet_name in sheet_names:
                            try:
                                sheet_df = pd.read_excel(str(input_path), sheet_name=sheet_name)
                                
                                # Debug: Print sheet info
                                safe_print(f"🔧 DEBUG: Analyzing sheet '{sheet_name}' with {len(sheet_df)} rows")
                                safe_print(f"🔧 DEBUG: Sheet columns: {list(sheet_df.columns)}")
                                safe_print(f"🔧 DEBUG: First few rows:\n{sheet_df.head()}")
                                
                                # Simple heuristic: BRIEF has fewer rows, DETAILED has more
                                if len(sheet_df) <= 15:
                                    sheet_type = 'BRIEF'
                                    
                                    # Try to extract BRIEF data from any sheet with reasonable data
                                    if len(sheet_df) >= 3 and len(sheet_df) <= 20:  # More flexible range
                                        brief_sheet_data = []
                                        
                                        safe_print(f"🔧 DEBUG: Attempting BRIEF extraction from sheet '{sheet_name}'")
                                        
                                        for idx, row in sheet_df.iterrows():
                                            # Extract BRIEF data for aspect summary
                                            brief_row = {}
                                            
                                            # More flexible column matching
                                            for col in sheet_df.columns:
                                                col_str = str(col).strip()
                                                col_lower = col_str.lower()
                                                
                                                # Match various column patterns
                                                if any(keyword in col_lower for keyword in ['aspek', 'section', 'aspect']):
                                                    brief_row['aspek'] = str(row[col]).strip() if pd.notna(row[col]) else ''
                                                elif any(keyword in col_lower for keyword in ['deskripsi', 'description', 'desc']):
                                                    brief_row['deskripsi'] = str(row[col]).strip() if pd.notna(row[col]) else ''
                                                elif any(keyword in col_lower for keyword in ['bobot', 'weight', 'berat']):
                                                    try:
                                                        brief_row['bobot'] = float(row[col]) if pd.notna(row[col]) else 0.0
                                                    except (ValueError, TypeError):
                                                        brief_row['bobot'] = 0.0
                                                elif any(keyword in col_lower for keyword in ['skor', 'score', 'nilai']):
                                                    try:
                                                        brief_row['skor'] = float(row[col]) if pd.notna(row[col]) else 0.0
                                                    except (ValueError, TypeError):
                                                        brief_row['skor'] = 0.0
                                                elif any(keyword in col_lower for keyword in ['capaian', 'achievement', 'pencapaian']):
                                                    try:
                                                        brief_row['capaian'] = float(row[col]) if pd.notna(row[col]) else 0.0
                                                    except (ValueError, TypeError):
                                                        brief_row['capaian'] = 0.0
                                                elif any(keyword in col_lower for keyword in ['penjelasan', 'explanation', 'keterangan']):
                                                    brief_row['penjelasan'] = str(row[col]).strip() if pd.notna(row[col]) else ''
                                            
                                            # Debug: show what we extracted for this row
                                            safe_print(f"🔧 DEBUG: Row {idx}: {brief_row}")
                                            
                                            # Add row if it has meaningful data (aspek is required)
                                            if brief_row.get('aspek') and brief_row.get('aspek').strip() and brief_row.get('aspek') != 'nan':
                                                brief_sheet_data.append(brief_row)
                                        
                                        safe_print(f"🔧 DEBUG: Successfully extracted {len(brief_sheet_data)} BRIEF summary rows from sheet '{sheet_name}'")
                                        
                                else:
                                    sheet_type = 'DETAILED'
                                    
                                sheet_analysis['sheet_types'][sheet_name] = {
                                    'type': sheet_type,
                                    'row_count': len(sheet_df),
                                    'contains_summary_data': len(sheet_df) <= 10 and len(sheet_df) >= 5
                                }
                            except Exception as e:
                                sheet_analysis['sheet_types'][sheet_name] = {
                                    'type': 'UNKNOWN',
                                    'error': str(e)
                                }
                        
                        extracted_data['sheet_analysis'] = sheet_analysis
                        extracted_data['brief_sheet_data'] = brief_sheet_data
                        
                    except Exception as e:
                        extracted_data['sheet_analysis'] = {
                            'error': f'Could not analyze sheets: {str(e)}'
                        }
                
            except Exception as read_error:
                extracted_data = {
                    'error': f'Could not read processed file: {str(read_error)}'
                }
        
        # Prepare response
        response_data = {
            'fileId': file_id,
            'originalFilename': original_filename,
            'processedFilename': output_filename,
            'fileType': file_type,
            'fileSize': input_path.stat().st_size,
            'uploadTime': datetime.now().isoformat(),
            'processing': processing_result,
            'extractedData': extracted_data,
            'metadata': {
                'checklistId': checklist_id,
                'year': year,
                'aspect': aspect
            }
        }
        
        return jsonify(response_data), 200
        
    except Exception as e:
        safe_print(f"🔧 DEBUG: Exception occurred: {str(e)}")
        import traceback
        safe_print(f"🔧 DEBUG: Full traceback: {traceback.format_exc()}")
        return jsonify({'error': f'Upload failed: {str(e)}'}), 500

@app.route('/api/download/<file_id>', methods=['GET'])
def download_file(file_id: str):
    """Download processed file by ID."""
    try:
        # Find the processed file
        for output_file in OUTPUT_FOLDER.glob(f"processed_{file_id}_*.xlsx"):
            if output_file.exists():
                return send_file(
                    str(output_file),
                    as_attachment=True,
                    download_name=f"GCG_Assessment_{file_id}.xlsx",
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
        
        return jsonify({'error': 'File not found'}), 404
        
    except Exception as e:
        return jsonify({'error': f'Download failed: {str(e)}'}), 500

@app.route('/api/files', methods=['GET'])
def list_files():
    """List all processed files."""
    try:
        files = []
        
        for output_file in OUTPUT_FOLDER.glob("processed_*.xlsx"):
            # Extract file ID from filename
            filename_parts = output_file.name.split('_', 2)
            if len(filename_parts) >= 2:
                file_id = filename_parts[1]
                
                # Get file stats
                stat = output_file.stat()
                
                files.append({
                    'fileId': file_id,
                    'filename': output_file.name,
                    'size': stat.st_size,
                    'created': datetime.fromtimestamp(stat.st_ctime).isoformat(),
                    'modified': datetime.fromtimestamp(stat.st_mtime).isoformat()
                })
        
        return jsonify({'files': files}), 200
        
    except Exception as e:
        return jsonify({'error': f'Failed to list files: {str(e)}'}), 500

@app.route('/api/system/info', methods=['GET'])
def system_info():
    """Get system information and capabilities."""
    return jsonify({
        'system': 'POS Data Cleaner 2',
        'version': '2.0.0',
        'capabilities': {
            'file_types': list(ALLOWED_EXTENSIONS),
            'formats_supported': ['DETAILED (56 rows)', 'BRIEF (13 rows)'],
            'languages': ['Indonesian'],
            'years_supported': '2014-2025',
            'gcg_aspects': ['I-VI (Roman numerals)', 'A-H (Alphabetic)', '1-10 (Numeric)'],
            'advanced_features': [
                'Mathematical topology processing',
                'Quantum superposition layouts', 
                'DNA helix patterns',
                'Fractal recursive structures',
                'Multi-engine OCR (Tesseract + PaddleOCR)',
                'ML classification (XGBoost + rules)'
            ]
        },
        'processing_pipeline': [
            'File type detection',
            'Format classification (DETAILED vs BRIEF)',
            'Pattern recognition (43+ indicator patterns)', 
            'Spatial matching (distance-based pairing)',
            'Manual.xlsx structure generation (362 rows)',
            'Quality validation'
        ],
        'infrastructure': {
            'privacy_first': True,
            'cloud_dependencies': None,
            'local_processing': True,
            'max_file_size': '16MB',
            'concurrent_processing': True
        }
    })


@app.route('/api/save', methods=['POST'])
def save_assessment():
    """
    Save assessment data directly to output.xlsx (no JSON intermediate)
    """
    try:
        data = request.json
        safe_print(f"🔧 DEBUG: Received save request with data keys: {data.keys()}")
        
        # Create assessment record
        assessment_id = f"{data.get('year', 'unknown')}_{data.get('auditor', 'unknown')}_{str(uuid.uuid4())[:8]}"
        saved_at = datetime.now().isoformat()
        
        # Load existing XLSX data and COMPLETELY REPLACE year's data (including deletions)
        all_rows = []
        existing_df = storage_service.read_excel('web-output/output.xlsx')
        
        if existing_df is not None:
            try:
                current_year = data.get('year')
                
                safe_print(f"🔧 DEBUG: Loading existing XLSX with {len(existing_df)} rows")
                safe_print(f"🔧 DEBUG: Current year to save: {current_year}")
                safe_print(f"🔧 DEBUG: Existing years in file: {existing_df['Tahun'].unique().tolist()}")
                
                # COMPLETELY REMOVE all existing data for this year (this handles deletions)
                if current_year:
                    original_count = len(existing_df)
                    existing_df = existing_df[existing_df['Tahun'] != current_year]
                    removed_count = original_count - len(existing_df)
                    safe_print(f"🔧 DEBUG: COMPLETELY REMOVED {removed_count} rows for year {current_year} (including deletions)")
                    safe_print(f"🔧 DEBUG: Preserved {len(existing_df)} rows from other years")
                
                # Convert remaining data back to list format
                for _, row in existing_df.iterrows():
                    all_rows.append(row.to_dict())
                    
                safe_print(f"🔧 DEBUG: Starting with {len(all_rows)} rows from other years")
            except Exception as e:
                safe_print(f"WARNING: Could not read existing XLSX: {e}")
        
        # Process new data and add to all_rows
        year = data.get('year', 'unknown')
        auditor = data.get('auditor', 'unknown')
        jenis_asesmen = data.get('jenis_asesmen', 'Internal')
        
        # Process main indicator data
        for row in data.get('data', []):
            # Map frontend data structure to XLSX format
            row_id = row.get('id', row.get('no', ''))
            section = row.get('aspek', row.get('section', ''))
            is_total = row.get('isTotal', False)
            
            # Determine Level and Type based on data structure
            if is_total:
                level = "1"
                row_type = "total"
            elif str(row_id).isdigit():
                level = "2"
                row_type = "indicator"
            else:
                level = "1"
                row_type = "header"
            
            xlsx_row = {
                'Level': level,
                'Type': row_type,
                'Section': section,
                'No': row_id,
                'Deskripsi': row.get('deskripsi', ''),
                'Bobot': row.get('bobot', ''),
                'Skor': row.get('skor', ''),
                'Capaian': row.get('capaian', ''),
                'Penjelasan': row.get('penjelasan', ''),
                'Tahun': year,
                'Penilai': auditor,
                'Jenis_Penilaian': jenis_asesmen,
                'Export_Date': saved_at[:10]
            }
            all_rows.append(xlsx_row)
        
        # Process aspect summary data (if provided)
        aspect_summary_data = data.get('aspectSummaryData', [])
        if aspect_summary_data:
            safe_print(f"🔧 DEBUG: Processing {len(aspect_summary_data)} aspect summary rows")
            
            for summary_row in aspect_summary_data:
                section = summary_row.get('aspek', '')
                deskripsi = summary_row.get('deskripsi', '')
                bobot = summary_row.get('bobot', 0)
                skor = summary_row.get('skor', 0)
                
                # Skip empty aspects or meaningless default data
                if not section or not deskripsi or (bobot == 0 and skor == 0):
                    continue
                    
                # Skip if this looks like an unedited default row (just roman numerals with no real data)
                if section in ['I', 'II', 'III', 'IV', 'V', 'VI'] and not deskripsi.strip():
                    continue
                    
                # Row 1: Header for this aspect
                header_row = {
                    'Level': "1",
                    'Type': 'header',
                    'Section': section,
                    'No': '',
                    'Deskripsi': summary_row.get('deskripsi', ''),
                    'Bobot': '',
                    'Skor': '',
                    'Capaian': '',
                    'Penjelasan': '',
                    'Tahun': year,
                    'Penilai': auditor,
                    'Jenis_Penilaian': jenis_asesmen,
                    'Export_Date': saved_at[:10]
                }
                all_rows.append(header_row)
                
                # Row 2: Subtotal for this aspect
                subtotal_row = {
                    'Level': "1", 
                    'Type': 'subtotal',
                    'Section': section,
                    'No': '',
                    'Deskripsi': f'JUMLAH {section}',
                    'Bobot': summary_row.get('bobot', ''),
                    'Skor': summary_row.get('skor', ''),
                    'Capaian': summary_row.get('capaian', ''),
                    'Penjelasan': summary_row.get('penjelasan', ''),
                    'Tahun': year,
                    'Penilai': auditor,
                    'Jenis_Penilaian': jenis_asesmen,
                    'Export_Date': saved_at[:10]
                }
                all_rows.append(subtotal_row)
        
        # Process separate totalData (total row sent separately from main data)
        total_data = data.get('totalData', {})
        if total_data and isinstance(total_data, dict):
            # Check if total data has meaningful values (not all zeros)
            has_meaningful_total = (
                total_data.get('bobot', 0) != 0 or 
                total_data.get('skor', 0) != 0 or 
                total_data.get('capaian', 0) != 0 or 
                total_data.get('penjelasan', '').strip() != ''
            )
            
            if has_meaningful_total:
                safe_print(f"🔧 DEBUG: Processing separate totalData: {total_data}")
                
                total_row = {
                    'Level': "4",
                    'Type': 'total',
                    'Section': 'TOTAL',
                    'No': '',
                    'Deskripsi': 'TOTAL',
                    'Bobot': total_data.get('bobot', ''),
                    'Skor': total_data.get('skor', ''),
                    'Capaian': total_data.get('capaian', ''),
                    'Penjelasan': total_data.get('penjelasan', ''),
                    'Tahun': year,
                    'Penilai': auditor,
                    'Jenis_Penilaian': jenis_asesmen,
                    'Export_Date': saved_at[:10]
                }
                all_rows.append(total_row)
                safe_print(f"🔧 DEBUG: Added totalData row to all_rows")
            else:
                safe_print(f"🔧 DEBUG: Skipping totalData - no meaningful values")
        
        # Convert to DataFrame and save XLSX
        if all_rows:
            df = pd.DataFrame(all_rows)
            
            # Remove any duplicate rows
            df_unique = df.drop_duplicates(subset=['Tahun', 'Section', 'No', 'Deskripsi'], keep='last')
            safe_print(f"🔧 DEBUG: Removed {len(df) - len(df_unique)} duplicate rows")
            
            # Custom sorting: year → aspek → no, then organize headers and subtotals properly
            def sort_key(row):
                # Ensure all values are consistently typed for comparison
                try:
                    year = int(row['Tahun']) if pd.notna(row['Tahun']) else 0
                except (ValueError, TypeError):
                    year = 0
                
                section = str(row['Section']) if pd.notna(row['Section']) else ''
                no = row['No']
                row_type = str(row['Type']) if pd.notna(row['Type']) else 'indicator'
                
                # Convert 'no' to numeric for proper sorting, handle empty values
                try:
                    no_numeric = int(no) if str(no).isdigit() else 9999
                except (ValueError, TypeError):
                    no_numeric = 9999
                
                # Type priority: header=0, indicators=1, subtotal=2, total=3 (appears last)
                type_priority = {'header': 0, 'indicator': 1, 'subtotal': 2, 'total': 3}.get(row_type, 1)
                
                # Special handling for total rows: they should appear at the very end of each year
                if row_type == 'total':
                    # Use 'ZZZZZ' as section to ensure total rows sort last within each year
                    section = 'ZZZZZ'
                
                return (year, section, type_priority, no_numeric)
            
            # Apply custom sorting
            df_sorted = df_unique.loc[df_unique.apply(sort_key, axis=1).sort_values().index]
            
            # Save XLSX using storage service
            success = storage_service.write_excel(df_sorted, 'web-output/output.xlsx')
            if success:
                safe_print(f"SUCCESS: Saved to output.xlsx with {len(df_sorted)} rows (sorted: year->aspek->no->type)")
            else:
                safe_print(f"ERROR: Failed to save output.xlsx")
            
        return jsonify({
            'success': True,
            'message': 'Data berhasil disimpan',
            'assessment_id': assessment_id,
            'saved_at': saved_at
        })
        
    except Exception as e:
        safe_print(f"ERROR: Error saving assessment: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


# generate_output_xlsx function removed - now saving directly to XLSX


@app.route('/api/delete-year-data', methods=['DELETE'])
def delete_year_data():
    """
    Delete all assessment data for a specific year from output.xlsx
    """
    try:
        data = request.json
        year_to_delete = data.get('year')
        
        if not year_to_delete:
            return jsonify({
                'success': False,
                'error': 'Year parameter is required'
            }), 400
        
        safe_print(f"🗑️ DEBUG: Received delete request for year: {year_to_delete}")
        
        # Load existing XLSX data
        existing_df = storage_service.read_excel('web-output/output.xlsx')
        
        if existing_df is None:
            return jsonify({
                'success': False,
                'error': 'No data file exists to delete from'
            }), 404
        
        try:
            safe_print(f"🔧 DEBUG: Loading existing XLSX with {len(existing_df)} rows")
            safe_print(f"🔧 DEBUG: Year to delete: {year_to_delete}")
            safe_print(f"🔧 DEBUG: Existing years in file: {existing_df['Tahun'].unique().tolist()}")
            
            # Check if the year exists in the data
            if year_to_delete not in existing_df['Tahun'].values:
                return jsonify({
                    'success': False,
                    'error': f'No data found for year {year_to_delete}'
                }), 404
            
            # Remove all data for the specified year
            original_count = len(existing_df)
            filtered_df = existing_df[existing_df['Tahun'] != year_to_delete]
            deleted_count = original_count - len(filtered_df)
            
            safe_print(f"🗑️ DEBUG: Deleted {deleted_count} rows for year {year_to_delete}")
            safe_print(f"🔧 DEBUG: Remaining {len(filtered_df)} rows from other years")
            
            # Save the filtered data back to the XLSX file
            if len(filtered_df) > 0:
                # Sort the remaining data properly before saving
                def sort_key(row):
                    year = row['Tahun']
                    section = str(row['Section']) if pd.notna(row['Section']) else ''
                    no = row['No']
                    row_type = str(row['Type']) if pd.notna(row['Type']) else 'indicator'
                    
                    # Convert 'no' to numeric for proper sorting
                    try:
                        no_numeric = int(no) if str(no).isdigit() else 9999
                    except (ValueError, TypeError):
                        no_numeric = 9999
                    
                    # Type priority: header=0, indicators=1, subtotal=2, total=3 (appears last)
                    type_priority = {'header': 0, 'indicator': 1, 'subtotal': 2, 'total': 3}.get(row_type, 1)
                    
                    # Special handling for total rows: they should appear at the very end of each year
                    if row_type == 'total':
                        # Use 'ZZZZZ' as section to ensure total rows sort last within each year
                        section = 'ZZZZZ'
                    
                    return (year, section, type_priority, no_numeric)
                
                # Apply sorting
                df_sorted = filtered_df.loc[filtered_df.apply(sort_key, axis=1).sort_values().index]
                success = storage_service.write_excel(df_sorted, 'web-output/output.xlsx')
                if success:
                    safe_print(f"SUCCESS: Updated output.xlsx with {len(df_sorted)} rows (deleted {deleted_count} rows for year {year_to_delete})")
                else:
                    safe_print(f"ERROR: Failed to update output.xlsx after deletion")
            else:
                # If no data remains, create an empty file with just headers
                empty_df = pd.DataFrame(columns=['Level', 'Type', 'Section', 'No', 'Deskripsi', 
                                               'Bobot', 'Skor', 'Capaian', 'Penjelasan', 'Tahun', 'Penilai', 
                                               'Jenis_Asesmen', 'Export_Date'])
                success = storage_service.write_excel(empty_df, 'web-output/output.xlsx')
                if success:
                    safe_print(f"SUCCESS: Created empty output.xlsx file (all data deleted)")
                else:
                    safe_print(f"ERROR: Failed to create empty output.xlsx file")
            
        except Exception as e:
            safe_print(f"ERROR: Could not process XLSX file: {e}")
            return jsonify({
                'success': False,
                'error': f'Could not process XLSX file: {str(e)}'
            }), 500
        
        return jsonify({
            'success': True,
            'message': f'Data untuk tahun {year_to_delete} berhasil dihapus',
            'deleted_rows': deleted_count,
            'year': year_to_delete
        })
        
    except Exception as e:
        safe_print(f"ERROR: Error deleting year data: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


@app.route('/api/load/<int:year>', methods=['GET'])
def load_assessment_by_year(year):
    """
    Load assessment data for a specific year from output.xlsx
    """
    try:
        # Read XLSX data
        df = storage_service.read_excel('web-output/output.xlsx')
        
        if df is None:
            return jsonify({
                'success': False,
                'data': [],
                'message': f'No saved data found for year {year}'
            })
        
        # Filter for the requested year
        year_df = df[df['Tahun'] == year]
        
        if len(year_df) > 0:
            safe_print(f"🔧 DEBUG: Processing {len(year_df)} rows for year {year}")
            
            # Detect format: BRIEF or DETAILED based on data types
            indicator_rows = year_df[year_df['Type'] == 'indicator']
            subtotal_rows = year_df[year_df['Type'] == 'subtotal'] 
            header_rows = year_df[year_df['Type'] == 'header']
            
            is_detailed = len(indicator_rows) > 10 and len(subtotal_rows) > 0
            format_type = 'DETAILED' if is_detailed else 'BRIEF'
            
            safe_print(f"🔧 DEBUG: Detected format: {format_type}")
            safe_print(f"🔧 DEBUG: Found {len(indicator_rows)} indicators, {len(subtotal_rows)} subtotals, {len(header_rows)} headers")
            
            # Process indicator data for main table (both BRIEF and DETAILED)
            main_table_data = []
            for _, row in indicator_rows.iterrows():
                row_id = row.get('No', '')
                if pd.isna(row_id) or str(row_id).lower() in ['nan', '', 'none']:
                    continue
                    
                aspek = str(row.get('Section', ''))
                deskripsi = str(row.get('Deskripsi', ''))
                if not aspek or not deskripsi:
                    continue
                
                penjelasan = row.get('Penjelasan', '')
                if pd.isna(penjelasan) or str(penjelasan).lower() == 'nan':
                    penjelasan = 'Tidak Baik'
                
                main_table_data.append({
                    'id': str(row_id),
                    'aspek': aspek,
                    'deskripsi': deskripsi,
                    'jumlah_parameter': int(row.get('Jumlah_Parameter', 0)) if pd.notna(row.get('Jumlah_Parameter')) else 0,
                    'bobot': float(row.get('Bobot', 0)) if pd.notna(row.get('Bobot')) else 0,
                    'skor': float(row.get('Skor', 0)) if pd.notna(row.get('Skor')) else 0,
                    'capaian': float(row.get('Capaian', 0)) if pd.notna(row.get('Capaian')) else 0,
                    'penjelasan': str(penjelasan)
                })
            
            # Process aspek summary data (subtotals) for DETAILED mode
            aspek_summary_data = []
            if is_detailed and len(subtotal_rows) > 0:
                for _, row in subtotal_rows.iterrows():
                    aspek = str(row.get('Section', ''))
                    if not aspek:
                        continue
                        
                    penjelasan = row.get('Penjelasan', '')
                    if pd.isna(penjelasan) or str(penjelasan).lower() == 'nan':
                        penjelasan = 'Tidak Baik'
                    
                    aspek_summary_data.append({
                        'id': f'summary-{aspek}',
                        'aspek': aspek,
                        'deskripsi': str(row.get('Deskripsi', '')),
                        'jumlah_parameter': int(row.get('Jumlah_Parameter', 0)) if pd.notna(row.get('Jumlah_Parameter')) else 0,
                        'bobot': float(row.get('Bobot', 0)) if pd.notna(row.get('Bobot')) else 0,
                        'skor': float(row.get('Skor', 0)) if pd.notna(row.get('Skor')) else 0,
                        'capaian': float(row.get('Capaian', 0)) if pd.notna(row.get('Capaian')) else 0,
                        'penjelasan': str(penjelasan)
                    })
            
            safe_print(f"🔧 DEBUG: Processed {len(main_table_data)} indicators, {len(aspek_summary_data)} aspect summaries")
            
            # Get auditor and jenis_asesmen from first row
            auditor = year_df.iloc[0].get('Penilai', 'Unknown') if len(year_df) > 0 else 'Unknown'
            jenis_asesmen = year_df.iloc[0].get('Jenis_Asesmen', 'Internal') if len(year_df) > 0 else 'Internal'
            
            return jsonify({
                'success': True,
                'data': main_table_data,
                'aspek_summary_data': aspek_summary_data,
                'format_type': format_type,
                'is_detailed': is_detailed,
                'auditor': auditor,
                'jenis_asesmen': jenis_asesmen,
                'method': 'xlsx_load',
                'saved_at': year_df.iloc[0].get('Export_Date', '') if len(year_df) > 0 else '',
                'message': f'Loaded {len(main_table_data)} indicators + {len(aspek_summary_data)} summaries for year {year} ({format_type} format)'
            })
        else:
            return jsonify({
                'success': False,
                'data': [],
                'message': f'No saved data found for year {year}'
            })
            
    except Exception as e:
        safe_print(f"ERROR: Error loading year {year}: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e),
            'data': []
        }), 500


@app.route('/api/dashboard-data', methods=['GET'])
def get_dashboard_data():
    """
    Get all assessment data from output.xlsx for dashboard visualization
    """
    try:
        # Read XLSX data
        df = storage_service.read_excel('web-output/output.xlsx')
        
        if df is None:
            return jsonify({
                'success': False,
                'data': [],
                'message': 'No dashboard data available. Please save some assessments first.'
            })
        
        safe_print(f"🔧 DEBUG: Dashboard loading {len(df)} rows from output.xlsx")
        safe_print(f"🔧 DEBUG: Years in file: {df['Tahun'].unique().tolist()}")
        safe_print(f"🔧 DEBUG: Sample rows: {df[['Tahun', 'Section', 'Skor']].head().to_dict('records')}")
        
        # Convert to dashboard format
        dashboard_data = []
        for _, row in df.iterrows():
            # Handle NaN values properly
            bobot = row.get('Bobot', 0)
            skor = row.get('Skor', 0)
            capaian = row.get('Capaian', 0)
            jumlah_param = row.get('Jumlah_Parameter', 0)
            
            # Convert NaN to 0 for numeric fields
            if pd.isna(bobot):
                bobot = 0
            if pd.isna(skor):
                skor = 0
            if pd.isna(capaian):
                capaian = 0
            if pd.isna(jumlah_param):
                jumlah_param = 0
                
            dashboard_item = {
                'id': str(row.get('No', '')),
                'aspek': str(row.get('Section', '')),
                'deskripsi': str(row.get('Deskripsi', '')),
                'jumlah_parameter': float(jumlah_param),
                'bobot': float(bobot),
                'skor': float(skor),
                'capaian': float(capaian),
                'penjelasan': str(row.get('Penjelasan', '')),
                'year': int(row.get('Tahun', 2022)),
                'auditor': str(row.get('Penilai', 'Unknown')),
                'jenis_asesmen': str(row.get('Jenis_Asesmen', 'Internal'))
            }
            dashboard_data.append(dashboard_item)
        
        # Group by year for multi-year support
        years_data = {}
        for item in dashboard_data:
            year = item['year']
            if year not in years_data:
                years_data[year] = {
                    'year': year,
                    'auditor': item['auditor'],
                    'jenis_asesmen': item['jenis_asesmen'],
                    'data': []
                }
            years_data[year]['data'].append(item)
        
        return jsonify({
            'success': True,
            'years_data': years_data,
            'total_rows': len(dashboard_data),
            'available_years': list(years_data.keys()),
            'message': f'Loaded dashboard data for {len(years_data)} year(s)'
        })
        
    except Exception as e:
        safe_print(f"ERROR: Error loading dashboard data: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e),
            'data': []
        }), 500


@app.route('/api/aspek-data', methods=['GET'])
def get_aspek_data():
    """
    Get hybrid data (subtotal + header) for aspek summary table
    """
    try:
        # Read XLSX data
        df = storage_service.read_excel('web-output/output.xlsx')
        
        if df is None:
            return jsonify({
                'success': False,
                'data': [],
                'message': 'No data available'
            })
        
        # Create hybrid data: subtotal numeric data + header descriptions
        subtotal_rows = df[df['Type'] == 'subtotal']
        header_rows = df[df['Type'] == 'header']
        
        # Convert to frontend format by combining subtotal + header data
        indicators = []
        for _, subtotal_row in subtotal_rows.iterrows():
            # Find matching header row by Section and Year
            matching_header = header_rows[
                (header_rows['Section'] == subtotal_row['Section']) & 
                (header_rows['Tahun'] == subtotal_row['Tahun'])
            ]
            
            # Use header description if found, otherwise subtotal description
            deskripsi = subtotal_row['Deskripsi']  # fallback
            if not matching_header.empty:
                deskripsi = matching_header.iloc[0]['Deskripsi']
            indicators.append({
                'id': str(subtotal_row.get('No', '')),
                'aspek': str(subtotal_row.get('Section', '')),
                'deskripsi': deskripsi,  # Use header description
                'jumlah_parameter': int(subtotal_row.get('Jumlah_Parameter', 0)) if pd.notna(subtotal_row.get('Jumlah_Parameter')) else 0,
                'bobot': float(subtotal_row.get('Bobot', 0)) if pd.notna(subtotal_row.get('Bobot')) else 0,
                'skor': float(subtotal_row.get('Skor', 0)) if pd.notna(subtotal_row.get('Skor')) else 0,
                'capaian': float(subtotal_row.get('Capaian', 0)) if pd.notna(subtotal_row.get('Capaian')) else 0,
                'penjelasan': str(subtotal_row.get('Penjelasan', 'Tidak Baik')),
                'tahun': int(subtotal_row.get('Tahun', 0)) if pd.notna(subtotal_row.get('Tahun')) else 0
            })
        
        return jsonify({
            'success': True,
            'data': indicators,
            'total': len(indicators),
            'message': f'Loaded {len(indicators)} aspek records'
        })
        
    except Exception as e:
        safe_print(f"ERROR: Error loading aspek data: {e}")
        return jsonify({
            'success': False,
            'error': str(e),
            'data': []
        }), 500

def _cleanup_orphaned_data_internal():
    """
    Internal helper function to clean up orphaned data (without returning HTTP response)
    """
    assessments_path = Path(__file__).parent.parent / 'web-output' / 'assessments.json'
    
    # Get years that exist in output.xlsx
    xlsx_years = set()
    df = storage_service.read_excel('web-output/output.xlsx')
    if df is not None:
        xlsx_years = set(df['Tahun'].unique())
    
    # Clean up assessments.json
    orphaned_count = 0
    if assessments_path.exists():
        with open(assessments_path, 'r') as f:
            assessments_data = json.load(f)
        
        # Filter out orphaned entries
        cleaned_assessments = []
        for assessment in assessments_data.get('assessments', []):
            year = assessment.get('year')
            if year in xlsx_years:
                cleaned_assessments.append(assessment)
            else:
                orphaned_count += 1
        
        # Save cleaned data if any changes were made
        if orphaned_count > 0:
            assessments_data['assessments'] = cleaned_assessments
            with open(assessments_path, 'w') as f:
                json.dump(assessments_data, f, indent=2)
            safe_print(f"🔄 Auto-cleaned {orphaned_count} orphaned entries")
    
    return orphaned_count

@app.route('/api/indicator-data', methods=['GET'])
def get_indicator_data():
    """
    Get pure indicator data for detailed bottom table
    """
    try:
        # Auto-cleanup orphaned data before proceeding
        try:
            _cleanup_orphaned_data_internal()
        except Exception as cleanup_error:
            safe_print(f"WARNING: Auto-cleanup failed: {cleanup_error}")
        
        # Read XLSX data
        df = storage_service.read_excel('web-output/output.xlsx')
        
        if df is None:
            return jsonify({
                'success': False,
                'data': [],
                'message': 'No data available'
            })
        
        # Filter only indicator rows
        indicator_rows = df[df['Type'] == 'indicator']
        
        # Convert to frontend format
        indicators = []
        for _, row in indicator_rows.iterrows():
            indicators.append({
                'id': str(row.get('No', '')),
                'aspek': str(row.get('Section', '')),
                'deskripsi': str(row.get('Deskripsi', '')),
                'jumlah_parameter': int(row.get('Jumlah_Parameter', 0)) if pd.notna(row.get('Jumlah_Parameter')) else 0,
                'bobot': float(row.get('Bobot', 0)) if pd.notna(row.get('Bobot')) else 0,
                'skor': float(row.get('Skor', 0)) if pd.notna(row.get('Skor')) else 0,
                'capaian': float(row.get('Capaian', 0)) if pd.notna(row.get('Capaian')) else 0,
                'penjelasan': str(row.get('Penjelasan', 'Tidak Baik')),
                'tahun': int(row.get('Tahun', 0)) if pd.notna(row.get('Tahun')) else 0
            })
        
        return jsonify({
            'success': True,
            'data': indicators,
            'total': len(indicators),
            'message': f'Loaded {len(indicators)} indicator records'
        })
        
    except Exception as e:
        safe_print(f"ERROR: Error loading indicator data: {e}")
        return jsonify({
            'success': False,
            'error': str(e),
            'data': []
        }), 500

@app.route('/api/gcg-chart-data', methods=['GET'])
def get_gcg_chart_data():
    """
    Get assessment data formatted for GCGChart component (graphics-2 format)
    Returns data with Level hierarchy as expected by processGCGData function
    """
    try:
        # Read XLSX data
        df = storage_service.read_excel('web-output/output.xlsx')
        
        if df is None:
            safe_print(f"WARNING: output.xlsx not found or empty")
            return jsonify({
                'success': True,
                'data': [],
                'message': 'No chart data available. Please save some assessments first.'
            })
        
        safe_print(f"INFO: GCG Chart Data: Loading {len(df)} rows from output.xlsx")
        
        # Convert to graphics-2 GCGData format
        gcg_data = []
        for _, row in df.iterrows():
            # Determine level based on row type
            level = 3  # Default to section level
            row_type = str(row.get('Type', '')).lower()
            
            if row_type == 'total':
                level = 4
            elif row_type == 'header':
                level = 1
            elif row_type == 'indicator':
                level = 2
            elif row_type == 'subtotal':
                level = 3
            
            # Handle NaN values
            tahun = int(row.get('Tahun', 2022))
            skor = float(row.get('Skor', 0)) if not pd.isna(row.get('Skor', 0)) else 0
            capaian = float(row.get('Capaian', 0)) if not pd.isna(row.get('Capaian', 0)) else 0
            bobot = float(row.get('Bobot', 0)) if not pd.isna(row.get('Bobot', 0)) else None
            jumlah_param = float(row.get('Jumlah_Parameter', 0)) if not pd.isna(row.get('Jumlah_Parameter', 0)) else None
            
            gcg_item = {
                'Tahun': tahun,
                'Skor': skor,
                'Level': level,
                'Section': str(row.get('Section', '')),
                'Capaian': capaian,
                'Bobot': bobot,
                'Jumlah_Parameter': jumlah_param,
                'Penjelasan': str(row.get('Penjelasan', '')),
                'Penilai': str(row.get('Penilai', 'Unknown')),
                'No': str(row.get('No', '')),
                'Deskripsi': str(row.get('Deskripsi', '')),
                'Jenis_Penilaian': str(row.get('Jenis_Penilaian', 'Data Kosong'))
            }
            gcg_data.append(gcg_item)
        
        return jsonify({
            'success': True,
            'data': gcg_data,
            'total_rows': len(gcg_data),
            'available_years': list(set([item['Tahun'] for item in gcg_data])),
            'message': f'Loaded GCG chart data: {len(gcg_data)} rows'
        })
        
    except Exception as e:
        safe_print(f"ERROR: Error loading GCG chart data: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e),
            'data': []
        }), 500


@app.route('/api/gcg-mapping', methods=['GET'])
def get_gcg_mapping():
    """
    Get GCG mapping data for autocomplete suggestions
    """
    try:
        # Path to GCG mapping CSV file
        gcg_mapping_path = Path(__file__).parent.parent / 'GCG_MAPPING.csv'
        
        if not gcg_mapping_path.exists():
            safe_print(f"WARNING: GCG_MAPPING.csv not found at: {gcg_mapping_path}")
            return jsonify({
                'success': False,
                'error': 'GCG mapping file not found',
                'data': []
            }), 404
        
        # Read GCG mapping CSV
        df = pd.read_csv(gcg_mapping_path)
        
        # Convert to list of dictionaries for JSON response
        gcg_data = []
        for _, row in df.iterrows():
            gcg_item = {
                'level': str(row.get('Level', '')),
                'type': str(row.get('Type', '')),
                'section': str(row.get('Section', '')),
                'no': str(row.get('No', '')),
                'deskripsi': str(row.get('Deskripsi', '')),
                'jumlah_parameter': str(row.get('Jumlah_Parameter', '')),
                'bobot': str(row.get('Bobot', ''))
            }
            gcg_data.append(gcg_item)
        
        # Return all items for flexible filtering on frontend
        return jsonify({
            'success': True,
            'data': gcg_data,
            'total_items': len(gcg_data),
            'headers': len([item for item in gcg_data if item['type'] == 'header']),
            'indicators': len([item for item in gcg_data if item['type'] == 'indicator']),
            'message': f'Loaded {len(gcg_data)} GCG items for autocomplete'
        })
        
    except Exception as e:
        safe_print(f"ERROR: Error loading GCG mapping: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e),
            'data': []
        }), 500

@app.route('/api/cleanup-orphaned-data', methods=['POST'])
def cleanup_orphaned_data():
    """
    Clean up orphaned entries in assessments.json that don't exist in output.xlsx
    """
    try:
        assessments_path = Path(__file__).parent.parent / 'web-output' / 'assessments.json'
        
        # Get years that exist in output.xlsx
        xlsx_years = set()
        df = storage_service.read_excel('web-output/output.xlsx')
        if df is not None:
            xlsx_years = set(df['Tahun'].unique())
            safe_print(f"INFO: Found years in output.xlsx: {sorted(xlsx_years)}")
        else:
            safe_print("WARNING: output.xlsx not found - will clean all assessments.json entries")
        
        # Clean up assessments.json
        orphaned_count = 0
        if assessments_path.exists():
            with open(assessments_path, 'r') as f:
                assessments_data = json.load(f)
            
            original_count = len(assessments_data.get('assessments', []))
            
            # Filter out orphaned entries (keep only years that exist in xlsx or if xlsx doesn't exist, keep none)
            cleaned_assessments = []
            for assessment in assessments_data.get('assessments', []):
                year = assessment.get('year')
                if year in xlsx_years:
                    cleaned_assessments.append(assessment)
                else:
                    orphaned_count += 1
                    safe_print(f"CLEANUP: Removing orphaned assessment for year {year}")
            
            # Save cleaned data
            assessments_data['assessments'] = cleaned_assessments
            with open(assessments_path, 'w') as f:
                json.dump(assessments_data, f, indent=2)
                
            safe_print(f"SUCCESS: Cleaned up {orphaned_count} orphaned entries from assessments.json")
            safe_print(f"INFO: Kept {len(cleaned_assessments)} valid entries")
        else:
            safe_print("WARNING: assessments.json not found - nothing to clean")
        
        return jsonify({
            'success': True,
            'message': f'Successfully cleaned up {orphaned_count} orphaned entries',
            'orphaned_count': orphaned_count,
            'xlsx_years': sorted(list(xlsx_years)),
            'xlsx_exists': storage_service.file_exists('web-output/output.xlsx'),
            'assessments_exists': assessments_path.exists()
        })
        
    except Exception as e:
        safe_print(f"ERROR: Error during cleanup: {e}")
        return jsonify({
            'success': False,
            'error': str(e),
            'message': 'Failed to cleanup orphaned data'
        }), 500


@app.route('/api/uploaded-files', methods=['GET'])
def get_uploaded_files():
    """Get all uploaded files from storage."""
    try:
        # Get year filter from query parameters
        year = request.args.get('year')
        
        # Read uploaded files data from storage
        files_data = storage_service.read_excel('uploaded-files.xlsx')
        
        if files_data is None:
            # Return empty list if no files exist yet
            return jsonify({'files': []}), 200
        
        # Convert DataFrame to list of dictionaries, replacing NaN with None
        files_data = files_data.fillna('')  # Replace NaN with empty strings
        files_list = files_data.to_dict('records')
        
        # Filter by year if provided
        if year:
            try:
                year_int = int(year)
                files_list = [f for f in files_list if f.get('year') == year_int]
            except ValueError:
                return jsonify({'error': 'Invalid year parameter'}), 400
        
        return jsonify({'files': files_list}), 200
        
    except Exception as e:
        safe_print(f"Error getting uploaded files: {e}")
        return jsonify({'error': f'Failed to get uploaded files: {str(e)}'}), 500

@app.route('/api/uploaded-files', methods=['POST'])
def create_uploaded_file():
    """Add a new uploaded file record to storage."""
    try:
        data = request.get_json()
        
        # Validate required fields
        required_fields = ['fileName', 'fileSize', 'year', 'uploadDate']
        for field in required_fields:
            if field not in data:
                return jsonify({'error': f'Missing required field: {field}'}), 400
        
        # Read existing files data
        try:
            files_data = storage_service.read_excel('uploaded-files.xlsx')
            if files_data is None:
                # Create new DataFrame if no data exists
                files_data = pd.DataFrame()
        except:
            files_data = pd.DataFrame()
        
        # Generate unique ID
        new_id = str(uuid.uuid4())
        
        # Create new file record
        new_file = {
            'id': new_id,
            'fileName': data['fileName'],
            'fileSize': data['fileSize'],
            'uploadDate': data['uploadDate'],
            'year': data['year'],
            'checklistId': data.get('checklistId'),
            'checklistDescription': data.get('checklistDescription'),
            'aspect': data.get('aspect', 'Tidak Diberikan Aspek'),
            'subdirektorat': data.get('subdirektorat'),
            'status': 'uploaded',
            'filePath': data.get('filePath'),
            'uploadedBy': data.get('uploadedBy', 'Unknown User'),
            'userDirektorat': data.get('userDirektorat', 'Unknown'),
            'userSubdirektorat': data.get('userSubdirektorat', 'Unknown'),
            'userDivisi': data.get('userDivisi', 'Unknown'),
            'userWhatsApp': data.get('userWhatsApp', ''),
            'userEmail': data.get('userEmail', '')
        }
        
        # Add to DataFrame
        new_row = pd.DataFrame([new_file])
        files_data = pd.concat([files_data, new_row], ignore_index=True)
        
        # Save to storage
        success = storage_service.write_excel(files_data, 'uploaded-files.xlsx')
        
        if success:
            return jsonify({'success': True, 'file': new_file}), 201
        else:
            return jsonify({'error': 'Failed to save file to storage'}), 500
            
    except Exception as e:
        safe_print(f"Error creating uploaded file: {e}")
        return jsonify({'error': f'Failed to create uploaded file: {str(e)}'}), 500

@app.route('/api/fix-uploaded-files-schema', methods=['POST'])
def fix_uploaded_files_schema():
    """Add missing user information columns to uploaded-files.xlsx"""
    try:
        # Read existing files data
        files_data = storage_service.read_excel('uploaded-files.xlsx')
        
        if files_data is None:
            return jsonify({'error': 'No files data found'}), 404
        
        # Check if user columns already exist
        missing_columns = []
        required_user_columns = ['uploadedBy', 'userRole', 'userDirektorat', 'userSubdirektorat', 'userDivisi', 'userWhatsApp', 'userEmail']
        
        for col in required_user_columns:
            if col not in files_data.columns:
                missing_columns.append(col)
        
        if not missing_columns:
            return jsonify({
                'success': True,
                'message': 'All user columns already exist',
                'columns': required_user_columns
            }), 200
        
        # Add missing columns with default values
        for col in missing_columns:
            if col == 'uploadedBy':
                files_data[col] = 'Unknown User'
            elif col == 'userRole':
                files_data[col] = 'user'  # Default role
            elif col in ['userWhatsApp', 'userEmail']:
                files_data[col] = ''  # Empty string for contact fields
            else:
                files_data[col] = 'Unknown'
        
        safe_print(f"📝 Adding missing user columns: {missing_columns}")
        
        # Save updated data
        success = storage_service.write_excel(files_data, 'uploaded-files.xlsx')
        
        if success:
            return jsonify({
                'success': True,
                'message': f'Added missing user columns: {missing_columns}',
                'addedColumns': missing_columns,
                'totalRecords': len(files_data)
            }), 200
        else:
            return jsonify({'error': 'Failed to save changes to storage'}), 500
            
    except Exception as e:
        safe_print(f"Error fixing uploaded files schema: {e}")
        return jsonify({'error': f'Failed to fix schema: {str(e)}'}), 500

@app.route('/api/uploaded-files/<file_id>', methods=['DELETE'])
@app.route('/api/delete-file/<file_id>', methods=['DELETE'])
def delete_uploaded_file(file_id):
    """Delete an uploaded file record and actual file from local storage."""
    try:
        safe_print(f"🗑️ DELETE request received for file_id: {file_id}")

        # Read existing files data
        files_data = storage_service.read_excel('uploaded-files.xlsx')

        if files_data is None:
            safe_print(f"❌ No files data found in uploaded-files.xlsx")
            return jsonify({'error': 'No files data found'}), 404

        safe_print(f"📊 Total records in database: {len(files_data)}")
        safe_print(f"🔍 Searching for file_id: {file_id}")
        safe_print(f"🔍 Sample IDs in database: {files_data['id'].head(3).tolist()}")

        # Find the file record to get the file path
        file_record = files_data[files_data['id'] == file_id]

        # If not found by UUID, check if it's a fallback ID format: file_{checklistId}
        if file_record.empty and file_id.startswith('file_'):
            try:
                checklist_id = int(file_id.replace('file_', ''))
                safe_print(f"🔍 Fallback ID detected, searching by checklistId: {checklist_id}")

                # Try to find by checklistId
                file_record = files_data[files_data['checklistId'] == checklist_id]

                if not file_record.empty:
                    safe_print(f"✅ Found record by checklistId: {checklist_id}")
                    # If multiple records (shouldn't happen), take the most recent
                    if len(file_record) > 1:
                        safe_print(f"⚠️ Multiple records found for checklistId {checklist_id}, using most recent")
                        file_record = file_record.sort_values('uploadDate', ascending=False).head(1)
                else:
                    safe_print(f"❌ No record found for checklistId: {checklist_id}")
            except ValueError:
                safe_print(f"❌ Invalid fallback ID format: {file_id}")

        if file_record.empty:
            safe_print(f"❌ File not found in database with id: {file_id}")
            safe_print(f"📋 Sample IDs in database:")
            for idx, row_id in enumerate(files_data['id'].tolist()[:10]):
                safe_print(f"  {idx+1}. {row_id}")
            return jsonify({'error': 'File not found'}), 404

        safe_print(f"✅ File record found: {file_record.iloc[0]['fileName']}")
        
        # Get the file path (check both localFilePath and filePath for compatibility)
        file_path = file_record.iloc[0].get('localFilePath') or file_record.iloc[0].get('filePath')

        # Delete the actual file from local storage if path exists
        if file_path:
            try:
                # For local storage mode, delete from local filesystem
                local_file_path = Path(__file__).parent.parent / 'data' / file_path
                safe_print(f"🔧 DEBUG: Attempting to delete local file: {local_file_path}")

                if local_file_path.exists():
                    local_file_path.unlink()
                    safe_print(f"🗑️ Deleted file from local storage: {local_file_path}")

                    # Also try to clean up empty parent directories
                    try:
                        parent_dir = local_file_path.parent
                        if parent_dir.exists() and not any(parent_dir.iterdir()):
                            parent_dir.rmdir()
                            safe_print(f"🗑️ Removed empty directory: {parent_dir}")
                    except Exception as dir_error:
                        safe_print(f"⚠️ Warning: Could not remove directory: {dir_error}")
                else:
                    safe_print(f"⚠️ Warning: File not found in local storage: {local_file_path}")

            except Exception as file_delete_error:
                safe_print(f"⚠️ Warning: Failed to delete file from storage: {file_delete_error}")
                # Continue with database record deletion even if file deletion fails
        
        # Remove the file record from SQLite database (primary storage)
        from database import get_db_connection
        try:
            with get_db_connection() as conn:
                cursor = conn.cursor()

                # Delete by UUID first
                cursor.execute("DELETE FROM uploaded_files WHERE id = ?", (file_id,))
                deleted_count = cursor.rowcount

                # If not found by UUID and it's a fallback ID, try by checklist_id
                if deleted_count == 0 and file_id.startswith('file_'):
                    try:
                        checklist_id = int(file_id.replace('file_', ''))
                        cursor.execute("DELETE FROM uploaded_files WHERE checklist_id = ?", (checklist_id,))
                        deleted_count = cursor.rowcount
                    except ValueError:
                        pass

                if deleted_count == 0:
                    return jsonify({'error': 'File record not found in database'}), 404

                conn.commit()
                safe_print(f"🔧 DEBUG: Deleted {deleted_count} record(s) from database")

        except Exception as db_error:
            safe_print(f"🔧 ERROR: Database deletion failed: {db_error}")
            return jsonify({'error': f'Failed to delete from database: {str(db_error)}'}), 500

        # Also remove from Excel file for backward compatibility
        try:
            initial_count = len(files_data)
            files_data = files_data[files_data['id'] != file_id]

            # If not removed and it's a fallback ID, try by checklistId
            if len(files_data) == initial_count and file_id.startswith('file_'):
                try:
                    checklist_id = int(file_id.replace('file_', ''))
                    files_data = files_data[files_data['checklistId'] != checklist_id]
                except ValueError:
                    pass

            storage_service.write_excel(files_data, 'uploaded-files.xlsx')
            safe_print(f"🔧 DEBUG: Also removed from Excel for backward compatibility")
        except Exception as excel_error:
            # Non-critical - Excel is just backup
            safe_print(f"🔧 WARNING: Could not remove from Excel (non-critical): {excel_error}")

        return jsonify({
            'success': True,
            'message': f'File deleted from both database and storage',
            'deletedFilePath': file_path
        }), 200
            
    except Exception as e:
        safe_print(f"Error deleting uploaded file: {e}")
        return jsonify({'error': f'Failed to delete uploaded file: {str(e)}'}), 500

@app.route('/api/download-file/<file_id>', methods=['GET'], endpoint='download_uploaded_file')
def download_uploaded_file(file_id):
    """Download a file from storage using its file ID."""
    try:
        # Read file metadata to get the storage path
        files_data = storage_service.read_excel('uploaded-files.xlsx')
        
        if files_data is None:
            return jsonify({'error': 'No files data found'}), 404
        
        # Find the file record
        file_record = files_data[files_data['id'] == file_id]
        
        if file_record.empty:
            return jsonify({'error': 'File not found'}), 404
        
        file_info = file_record.iloc[0]
        file_path = file_info.get('filePath')
        filename = file_info.get('fileName', 'download')
        
        # Try multiple path options
        file_paths_to_try = []
        
        # Option 1: localFilePath (preferred)
        if file_info.get('localFilePath'):
            file_paths_to_try.append(file_info.get('localFilePath'))
        
        # Option 2: filePath
        if file_info.get('filePath'):
            file_paths_to_try.append(file_info.get('filePath'))
        
        # Option 3: Construct path from file info (with secure_filename)
        if file_info.get('year') and file_info.get('subdirektorat') and file_info.get('checklistId'):
            constructed_path = f"gcg-documents/{file_info['year']}/{secure_filename(file_info['subdirektorat'])}/{file_info['checklistId']}/{secure_filename(filename)}"
            file_paths_to_try.append(constructed_path)
        
        # Option 4: Construct path from file info (without secure_filename)
        if file_info.get('year') and file_info.get('subdirektorat') and file_info.get('checklistId'):
            constructed_path_original = f"gcg-documents/{file_info['year']}/{file_info['subdirektorat']}/{file_info['checklistId']}/{filename}"
            file_paths_to_try.append(constructed_path_original)
        
        
        if not file_paths_to_try:
            return jsonify({'error': 'File path not found in database'}), 404
        
        # Try to find the file using multiple paths
        file_found = False
        actual_file_path = None
        
        for file_path in file_paths_to_try:
            # For local storage, construct the full path
            local_file_path = Path(__file__).parent.parent / 'data' / file_path
            if local_file_path.exists():
                actual_file_path = local_file_path
                file_found = True
                break
        
        if not file_found:
            return jsonify({'error': 'File not found in storage'}), 404
        
        # Send the file for download
        return send_file(
            str(actual_file_path), 
            as_attachment=True, 
            download_name=filename,
            mimetype='application/octet-stream'
        )
        
    except Exception as e:
        safe_print(f"Error downloading file: {e}")
        return jsonify({'error': f'Failed to download file: {str(e)}'}), 500

@app.route('/api/files/<file_id>/view', methods=['GET'])
def view_file(file_id):
    """Get a public view URL for a processed file."""
    try:
        # Find the processed file in OUTPUT_FOLDER
        file_path = None
        filename = None
        
        for output_file in OUTPUT_FOLDER.glob("processed_*.xlsx"):
            # Extract file ID from filename
            filename_parts = output_file.name.split('_', 2)
            if len(filename_parts) >= 2 and filename_parts[1] == file_id:
                file_path = output_file
                filename = output_file.name
                break
        
        if not file_path or not file_path.exists():
            return jsonify({'success': False, 'error': 'File not found'}), 404
        
        # For processed files, return a download URL since we can't "view" Excel files in browser
        return jsonify({
            'success': True,
            'url': f'http://localhost:5000/api/files/{file_id}/download',
            'filename': filename
        }), 200
        
    except Exception as e:
        safe_print(f"Error viewing file: {e}")
        return jsonify({'success': False, 'error': f'Failed to view file: {str(e)}'}), 500

@app.route('/api/files/<file_id>/download', methods=['GET'])
def download_file_by_id(file_id):
    """Download a file using its file ID - supports both processed files and uploaded files."""
    try:
        safe_print(f"📥 Download request for file_id: {file_id}")

        # First, try to find in uploaded-files.xlsx
        try:
            files_data = storage_service.read_excel('uploaded-files.xlsx')
            if files_data is not None and not files_data.empty:
                # Find file by ID
                file_record = files_data[files_data['id'] == file_id]
                if not file_record.empty:
                    file_info = file_record.iloc[0]
                    local_file_path = file_info.get('localFilePath', '')
                    file_name = file_info.get('fileName', 'document')

                    safe_print(f"📂 Found in uploaded-files.xlsx: {file_name}")
                    safe_print(f"📂 Local path: {local_file_path}")

                    # Construct full path
                    full_path = Path(__file__).parent.parent / 'data' / local_file_path

                    safe_print(f"📂 Full path: {full_path}")

                    if full_path.exists():
                        safe_print(f"✅ File exists, sending for download")

                        # Detect MIME type
                        import mimetypes
                        mime_type, _ = mimetypes.guess_type(file_name)
                        if not mime_type:
                            mime_type = 'application/octet-stream'

                        return send_file(
                            str(full_path),
                            as_attachment=True,
                            download_name=file_name,
                            mimetype=mime_type
                        )
                    else:
                        safe_print(f"❌ File not found at path: {full_path}")
        except Exception as e:
            safe_print(f"⚠️ Error checking uploaded-files.xlsx: {e}")

        # Fallback: try to find processed file in OUTPUT_FOLDER
        file_path = None
        filename = None

        for output_file in OUTPUT_FOLDER.glob("processed_*.xlsx"):
            # Extract file ID from filename
            filename_parts = output_file.name.split('_', 2)
            if len(filename_parts) >= 2 and filename_parts[1] == file_id:
                file_path = output_file
                filename = output_file.name
                break

        if file_path and file_path.exists():
            safe_print(f"✅ Found processed file: {filename}")
            # Send the file for download
            return send_file(
                str(file_path),
                as_attachment=True,
                download_name=filename,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

        # File not found in either location
        safe_print(f"❌ File not found: {file_id}")
        return jsonify({'error': 'File not found'}), 404

    except Exception as e:
        safe_print(f"❌ Error downloading file: {e}")
        import traceback
        safe_print(f"❌ Traceback: {traceback.format_exc()}")
        return jsonify({'error': f'Failed to download file: {str(e)}'}), 500

# AOI TABLES ENDPOINTS
@app.route('/api/aoiTables', methods=['GET'])
def get_aoi_tables():
    """Get all AOI tables (filtered by user structure if applicable)"""
    try:
        # Get filter parameters from query string
        user_role = request.args.get('userRole', '')
        user_subdirektorat = request.args.get('userSubdirektorat', '')
        user_divisi = request.args.get('userDivisi', '')
        year = request.args.get('year', '')

        aoi_data = storage_service.read_csv('config/aoi-tables.csv')
        if aoi_data is not None:
            # Replace NaN values with empty strings before converting to dict
            aoi_data = aoi_data.fillna('')

            # Filter by year if provided
            if year:
                aoi_data = aoi_data[aoi_data['tahun'] == int(year)]

            # Filter by user structure (unless super-admin)
            if user_role != 'super-admin' and user_subdirektorat:
                # Filter logic: match by divisi (most specific), subdirektorat, or direktorat
                filtered_tables = []
                for _, table in aoi_data.iterrows():
                    # Match by divisi (most specific)
                    if table.get('targetDivisi') and table['targetDivisi'] != 'Tidak ada':
                        if table['targetDivisi'] == user_divisi:
                            filtered_tables.append(table)
                    # Match by subdirektorat
                    elif table.get('targetSubdirektorat') and table['targetSubdirektorat'] != 'Tidak ada':
                        if table['targetSubdirektorat'] == user_subdirektorat:
                            filtered_tables.append(table)
                    # Match by direktorat - would need struktur data to check
                    # For now, we rely on frontend filtering for direktorat level
                    elif table.get('targetDirektorat') and table['targetDirektorat'] != 'Tidak ada':
                        # Include direktorat-level tables for now
                        filtered_tables.append(table)

                if filtered_tables:
                    aoi_data = pd.DataFrame(filtered_tables)
                else:
                    return jsonify([]), 200

            aoi_tables = aoi_data.to_dict(orient='records')
            return jsonify(aoi_tables), 200
        return jsonify([]), 200
    except Exception as e:
        safe_print(f"Error getting AOI tables: {e}")
        return jsonify([]), 200

@app.route('/api/aoiTables/<int:table_id>', methods=['GET'])
def get_aoi_table_by_id(table_id):
    """Get AOI table by ID"""
    try:
        aoi_data = storage_service.read_csv('config/aoi-tables.csv')
        if aoi_data is not None:
            aoi_data = aoi_data.fillna('')
            table_row = aoi_data[aoi_data['id'] == table_id]
            if not table_row.empty:
                table = safe_serialize_dict(table_row.iloc[0].to_dict())
                return jsonify(table), 200
        return jsonify({'error': 'AOI table not found'}), 404
    except Exception as e:
        safe_print(f"Error getting AOI table {table_id}: {e}")
        return jsonify({'error': f'Failed to get AOI table: {str(e)}'}), 500

@app.route('/api/aoiTables', methods=['POST'])
def create_aoi_table():
    """Create a new AOI table"""
    try:
        data = request.get_json()
        
        # Generate unique ID
        table_id = generate_unique_id()
        
        # Create AOI table object
        aoi_table_data = {
            'id': table_id,
            'nama': data.get('nama', ''),
            'tahun': data.get('tahun'),
            'targetType': data.get('targetType', ''),
            'targetDirektorat': data.get('targetDirektorat', ''),
            'targetSubdirektorat': data.get('targetSubdirektorat', ''),
            'targetDivisi': data.get('targetDivisi', ''),
            'createdAt': data.get('createdAt', datetime.now().isoformat()),
            'status': data.get('status', 'active')
        }
        
        # Read existing AOI tables
        existing_data = storage_service.read_csv('config/aoi-tables.csv')
        if existing_data is not None:
            aoi_df = existing_data
            # Ensure proper data types for string columns
            string_columns = ['nama', 'targetType', 'targetDirektorat', 'targetSubdirektorat', 'targetDivisi', 'status']
            for col in string_columns:
                if col in aoi_df.columns:
                    aoi_df[col] = aoi_df[col].astype(str).replace('nan', '')
        else:
            aoi_df = pd.DataFrame()

        # Add new AOI table
        new_aoi_df = pd.DataFrame([aoi_table_data])
        updated_df = pd.concat([aoi_df, new_aoi_df], ignore_index=True)

        # Ensure all string columns have proper dtype
        string_columns = ['nama', 'targetType', 'targetDirektorat', 'targetSubdirektorat', 'targetDivisi', 'status']
        for col in string_columns:
            if col in updated_df.columns:
                updated_df[col] = updated_df[col].astype(str).replace('nan', '')
        
        # Save to storage
        success = storage_service.write_csv(updated_df, 'config/aoi-tables.csv')
        
        if success:
            return jsonify(aoi_table_data), 201
        else:
            return jsonify({'error': 'Failed to save AOI table'}), 500
            
    except Exception as e:
        safe_print(f"Error creating AOI table: {e}")
        return jsonify({'error': f'Failed to create AOI table: {str(e)}'}), 500

@app.route('/api/aoiTables/<int:table_id>', methods=['PUT'])
def update_aoi_table(table_id):
    """Update an existing AOI table"""
    try:
        data = request.get_json()
        
        # Read existing AOI tables
        aoi_data = storage_service.read_csv('config/aoi-tables.csv')
        if aoi_data is None:
            return jsonify({'error': 'No AOI tables found'}), 404

        # Ensure proper data types for string columns
        string_columns = ['nama', 'targetType', 'targetDirektorat', 'targetSubdirektorat', 'targetDivisi', 'status']
        for col in string_columns:
            if col in aoi_data.columns:
                aoi_data[col] = aoi_data[col].astype(str).replace('nan', '')

        # Find the row to update
        mask = aoi_data['id'] == table_id
        if not mask.any():
            return jsonify({'error': 'AOI table not found'}), 404

        # Update the AOI table using proper assignment
        aoi_data.loc[mask, 'nama'] = str(data.get('nama', ''))
        aoi_data.loc[mask, 'tahun'] = int(data.get('tahun')) if data.get('tahun') else None
        aoi_data.loc[mask, 'targetType'] = str(data.get('targetType', ''))
        aoi_data.loc[mask, 'targetDirektorat'] = str(data.get('targetDirektorat', ''))
        aoi_data.loc[mask, 'targetSubdirektorat'] = str(data.get('targetSubdirektorat', ''))
        aoi_data.loc[mask, 'targetDivisi'] = str(data.get('targetDivisi', ''))
        aoi_data.loc[mask, 'status'] = str(data.get('status', 'active'))
        
        # Save to storage
        success = storage_service.write_csv(aoi_data, 'config/aoi-tables.csv')
        
        if success:
            # Return updated table
            updated_row = aoi_data[aoi_data['id'] == table_id]
            if not updated_row.empty:
                updated_table = safe_serialize_dict(updated_row.iloc[0].to_dict())
                return jsonify(updated_table), 200
            else:
                return jsonify({'error': 'AOI table not found after update'}), 404
        else:
            return jsonify({'error': 'Failed to update AOI table'}), 500
            
    except Exception as e:
        safe_print(f"Error updating AOI table {table_id}: {e}")
        return jsonify({'error': f'Failed to update AOI table: {str(e)}'}), 500

@app.route('/api/aoiTables/<int:table_id>', methods=['DELETE'])
def delete_aoi_table(table_id):
    """Delete an AOI table"""
    try:
        # Read existing AOI tables
        aoi_data = storage_service.read_csv('config/aoi-tables.csv')
        if aoi_data is None:
            return jsonify({'error': 'No AOI tables found'}), 404
        
        # Remove the AOI table
        aoi_data = aoi_data[aoi_data['id'] != table_id]
        
        # Save to storage
        success = storage_service.write_csv(aoi_data, 'config/aoi-tables.csv')
        
        if success:
            return jsonify({'message': f'AOI table {table_id} deleted successfully'}), 200
        else:
            return jsonify({'error': 'Failed to delete AOI table'}), 500
            
    except Exception as e:
        safe_print(f"Error deleting AOI table {table_id}: {e}")
        return jsonify({'error': f'Failed to delete AOI table: {str(e)}'}), 500

# AOI RECOMMENDATIONS ENDPOINTS
@app.route('/api/aoiRecommendations', methods=['GET'])
def get_aoi_recommendations():
    """Get AOI recommendations, optionally filtered by aoiTableId"""
    try:
        aoi_table_id = request.args.get('aoiTableId', type=int)
        
        recommendations_data = storage_service.read_csv('config/aoi-recommendations.csv')
        if recommendations_data is not None:
            recommendations_data = recommendations_data.fillna('')
            
            # Filter by aoiTableId if provided
            if aoi_table_id:
                recommendations_data = recommendations_data[recommendations_data['aoiTableId'] == aoi_table_id]
            
            recommendations = recommendations_data.to_dict(orient='records')
            return jsonify(recommendations), 200
        return jsonify([]), 200
    except Exception as e:
        safe_print(f"Error getting AOI recommendations: {e}")
        return jsonify([]), 200

@app.route('/api/aoiRecommendations/<int:recommendation_id>', methods=['GET'])
def get_aoi_recommendation_by_id(recommendation_id):
    """Get AOI recommendation by ID"""
    try:
        recommendations_data = storage_service.read_csv('config/aoi-recommendations.csv')
        if recommendations_data is not None:
            recommendations_data = recommendations_data.fillna('')
            recommendation_row = recommendations_data[recommendations_data['id'] == recommendation_id]
            if not recommendation_row.empty:
                recommendation = safe_serialize_dict(recommendation_row.iloc[0].to_dict())
                return jsonify(recommendation), 200
        return jsonify({'error': 'AOI recommendation not found'}), 404
    except Exception as e:
        safe_print(f"Error getting AOI recommendation {recommendation_id}: {e}")
        return jsonify({'error': f'Failed to get AOI recommendation: {str(e)}'}), 500

@app.route('/api/aoiRecommendations', methods=['POST'])
def create_aoi_recommendation():
    """Create a new AOI recommendation"""
    try:
        data = request.get_json()
        
        # Generate unique ID
        recommendation_id = generate_unique_id()
        
        # Read existing AOI recommendations
        existing_data = storage_service.read_csv('config/aoi-recommendations.csv')
        if existing_data is not None:
            recommendations_df = existing_data
        else:
            recommendations_df = pd.DataFrame()
        
        # Calculate correct row number for this table
        table_id = data.get('aoiTableId')
        if table_id and not recommendations_df.empty:
            # Get existing recommendations for this table
            table_recs = recommendations_df[recommendations_df['aoiTableId'] == table_id]
            if not table_recs.empty:
                # Since we're using sequential numbering (no gaps), just count existing + 1
                next_no = len(table_recs) + 1
            else:
                next_no = 1
        else:
            next_no = 1
        
        # Create AOI recommendation object
        aoi_recommendation_data = {
            'id': recommendation_id,
            'aoiTableId': table_id,
            'jenis': data.get('jenis', 'REKOMENDASI'),
            'no': next_no,  # Use calculated number instead of frontend value
            'isi': data.get('isi', ''),
            'tingkatUrgensi': data.get('tingkatUrgensi', 'SEDANG'),
            'aspekAOI': data.get('aspekAOI', ''),
            'pihakTerkait': data.get('pihakTerkait', ''),
            'organPerusahaan': data.get('organPerusahaan', ''),
            'createdAt': data.get('createdAt', datetime.now().isoformat()),
            'status': data.get('status', 'active')
        }
        
        # Add new AOI recommendation
        new_recommendation_df = pd.DataFrame([aoi_recommendation_data])
        updated_df = pd.concat([recommendations_df, new_recommendation_df], ignore_index=True)
        
        # Save to storage
        success = storage_service.write_csv(updated_df, 'config/aoi-recommendations.csv')
        
        if success:
            return jsonify(aoi_recommendation_data), 201
        else:
            return jsonify({'error': 'Failed to save AOI recommendation'}), 500
            
    except Exception as e:
        safe_print(f"Error creating AOI recommendation: {e}")
        return jsonify({'error': f'Failed to create AOI recommendation: {str(e)}'}), 500

@app.route('/api/aoiRecommendations/<int:recommendation_id>', methods=['PUT'])
def update_aoi_recommendation(recommendation_id):
    """Update an existing AOI recommendation"""
    try:
        data = request.get_json()
        
        # Read existing AOI recommendations
        recommendations_data = storage_service.read_csv('config/aoi-recommendations.csv')
        if recommendations_data is None:
            return jsonify({'error': 'No AOI recommendations found'}), 404
        
        # Update the AOI recommendation
        recommendations_data.loc[recommendations_data['id'] == recommendation_id, 'aoiTableId'] = data.get('aoiTableId')
        recommendations_data.loc[recommendations_data['id'] == recommendation_id, 'jenis'] = data.get('jenis', 'REKOMENDASI')
        recommendations_data.loc[recommendations_data['id'] == recommendation_id, 'no'] = data.get('no', 1)
        recommendations_data.loc[recommendations_data['id'] == recommendation_id, 'isi'] = data.get('isi', '')
        recommendations_data.loc[recommendations_data['id'] == recommendation_id, 'tingkatUrgensi'] = data.get('tingkatUrgensi', 'SEDANG')
        recommendations_data.loc[recommendations_data['id'] == recommendation_id, 'aspekAOI'] = data.get('aspekAOI', '')
        recommendations_data.loc[recommendations_data['id'] == recommendation_id, 'pihakTerkait'] = data.get('pihakTerkait', '')
        recommendations_data.loc[recommendations_data['id'] == recommendation_id, 'organPerusahaan'] = data.get('organPerusahaan', '')
        recommendations_data.loc[recommendations_data['id'] == recommendation_id, 'status'] = data.get('status', 'active')
        
        # Save to storage
        success = storage_service.write_csv(recommendations_data, 'config/aoi-recommendations.csv')
        
        if success:
            # Return updated recommendation
            updated_row = recommendations_data[recommendations_data['id'] == recommendation_id]
            if not updated_row.empty:
                updated_recommendation = safe_serialize_dict(updated_row.iloc[0].to_dict())
                return jsonify(updated_recommendation), 200
            else:
                return jsonify({'error': 'AOI recommendation not found after update'}), 404
        else:
            return jsonify({'error': 'Failed to update AOI recommendation'}), 500
            
    except Exception as e:
        safe_print(f"Error updating AOI recommendation {recommendation_id}: {e}")
        return jsonify({'error': f'Failed to update AOI recommendation: {str(e)}'}), 500

@app.route('/api/aoiRecommendations/<int:recommendation_id>', methods=['DELETE'])
def delete_aoi_recommendation(recommendation_id):
    """Delete an AOI recommendation and renumber remaining recommendations"""
    try:
        # Read existing AOI recommendations
        recommendations_data = storage_service.read_csv('config/aoi-recommendations.csv')
        if recommendations_data is None:
            return jsonify({'error': 'No AOI recommendations found'}), 404
        
        # Find the recommendation to be deleted to get its table ID and current number
        target_rec = recommendations_data[recommendations_data['id'] == recommendation_id]
        if target_rec.empty:
            return jsonify({'error': 'AOI recommendation not found'}), 404
        
        target_table_id = target_rec.iloc[0]['aoiTableId']
        target_no = int(target_rec.iloc[0]['no'])
        
        # Remove the target recommendation
        recommendations_data = recommendations_data[recommendations_data['id'] != recommendation_id]
        
        # Get all recommendations for the same table, sorted by 'no'
        table_recs = recommendations_data[recommendations_data['aoiTableId'] == target_table_id].copy()
        
        if not table_recs.empty:
            # Renumber all recommendations with 'no' greater than the deleted one
            # Sort by 'no' to ensure proper order
            table_recs_sorted = table_recs.sort_values('no')
            
            # Update the row numbers: shift down all numbers greater than target_no
            for idx, row in table_recs_sorted.iterrows():
                current_no = int(row['no'])
                if current_no > target_no:
                    recommendations_data.loc[idx, 'no'] = current_no - 1
        
        # Save to storage
        success = storage_service.write_csv(recommendations_data, 'config/aoi-recommendations.csv')
        
        if success:
            return jsonify({
                'message': f'AOI recommendation {recommendation_id} deleted successfully',
                'renumbered': f'Renumbered recommendations for table {target_table_id}'
            }), 200
        else:
            return jsonify({'error': 'Failed to delete AOI recommendation'}), 500
            
    except Exception as e:
        safe_print(f"Error deleting AOI recommendation {recommendation_id}: {e}")
        return jsonify({'error': f'Failed to delete AOI recommendation: {str(e)}'}), 500

# AOI DOCUMENTS ENDPOINTS
@app.route('/api/aoiDocuments', methods=['GET'])
def get_aoi_documents():
    """Get AOI documents, optionally filtered by recommendation ID or year"""
    try:
        aoi_recommendation_id = request.args.get('aoiRecommendationId', type=int)
        tahun = request.args.get('tahun', type=int)
        
        documents_data = storage_service.read_csv('config/aoi-documents.csv')
        if documents_data is not None:
            documents_data = documents_data.fillna('')
            
            # Filter by aoiRecommendationId if provided
            if aoi_recommendation_id:
                documents_data = documents_data[documents_data['aoiRecommendationId'] == aoi_recommendation_id]
            
            # Filter by tahun if provided
            if tahun:
                documents_data = documents_data[documents_data['tahun'] == tahun]
            
            documents = documents_data.to_dict(orient='records')
            return jsonify(documents), 200
        return jsonify([]), 200
    except Exception as e:
        safe_print(f"Error getting AOI documents: {e}")
        return jsonify([]), 200

@app.route('/api/aoiDocuments/<string:document_id>', methods=['GET'])
def get_aoi_document_by_id(document_id):
    """Get AOI document by ID"""
    try:
        documents_data = storage_service.read_csv('config/aoi-documents.csv')
        if documents_data is not None:
            documents_data = documents_data.fillna('')
            document_row = documents_data[documents_data['id'] == document_id]
            if not document_row.empty:
                document = safe_serialize_dict(document_row.iloc[0].to_dict())
                return jsonify(document), 200
        return jsonify({'error': 'AOI document not found'}), 404
    except Exception as e:
        safe_print(f"Error getting AOI document {document_id}: {e}")
        return jsonify({'error': f'Failed to get AOI document: {str(e)}'}), 500

@app.route('/api/aoiDocuments', methods=['POST'])
def create_aoi_document():
    """Create a new AOI document record"""
    try:
        data = request.get_json()
        
        # Generate unique ID (using string for AOI documents)
        document_id = f"aoi_{generate_unique_id()}"
        
        # Create AOI document object
        aoi_document_data = {
            'id': document_id,
            'fileName': data.get('fileName', ''),
            'fileSize': data.get('fileSize', 0),
            'uploadDate': data.get('uploadDate', datetime.now().isoformat()),
            'aoiRecommendationId': data.get('aoiRecommendationId'),
            'aoiJenis': data.get('aoiJenis', 'REKOMENDASI'),
            'aoiUrutan': data.get('aoiUrutan', 1),
            'userId': data.get('userId', ''),
            'userDirektorat': data.get('userDirektorat', ''),
            'userSubdirektorat': data.get('userSubdirektorat', ''),
            'userDivisi': data.get('userDivisi', ''),
            'fileType': data.get('fileType', ''),
            'status': data.get('status', 'active'),
            'tahun': data.get('tahun')
        }
        
        # Read existing AOI documents
        existing_data = storage_service.read_csv('config/aoi-documents.csv')
        if existing_data is not None:
            documents_df = existing_data
        else:
            documents_df = pd.DataFrame()
        
        # Add new AOI document
        new_document_df = pd.DataFrame([aoi_document_data])
        updated_df = pd.concat([documents_df, new_document_df], ignore_index=True)
        
        # Save to storage
        success = storage_service.write_csv(updated_df, 'config/aoi-documents.csv')
        
        if success:
            return jsonify(aoi_document_data), 201
        else:
            return jsonify({'error': 'Failed to save AOI document'}), 500
            
    except Exception as e:
        safe_print(f"Error creating AOI document: {e}")
        return jsonify({'error': f'Failed to create AOI document: {str(e)}'}), 500

@app.route('/api/aoiDocuments/<string:document_id>', methods=['PUT'])
def update_aoi_document(document_id):
    """Update an existing AOI document"""
    try:
        data = request.get_json()
        
        # Read existing AOI documents
        documents_data = storage_service.read_csv('config/aoi-documents.csv')
        if documents_data is None:
            return jsonify({'error': 'No AOI documents found'}), 404
        
        # Update the AOI document
        documents_data.loc[documents_data['id'] == document_id, 'fileName'] = data.get('fileName', '')
        documents_data.loc[documents_data['id'] == document_id, 'fileSize'] = data.get('fileSize', 0)
        documents_data.loc[documents_data['id'] == document_id, 'aoiRecommendationId'] = data.get('aoiRecommendationId')
        documents_data.loc[documents_data['id'] == document_id, 'aoiJenis'] = data.get('aoiJenis', 'REKOMENDASI')
        documents_data.loc[documents_data['id'] == document_id, 'aoiUrutan'] = data.get('aoiUrutan', 1)
        documents_data.loc[documents_data['id'] == document_id, 'userId'] = data.get('userId', '')
        documents_data.loc[documents_data['id'] == document_id, 'userDirektorat'] = data.get('userDirektorat', '')
        documents_data.loc[documents_data['id'] == document_id, 'userSubdirektorat'] = data.get('userSubdirektorat', '')
        documents_data.loc[documents_data['id'] == document_id, 'userDivisi'] = data.get('userDivisi', '')
        documents_data.loc[documents_data['id'] == document_id, 'fileType'] = data.get('fileType', '')
        documents_data.loc[documents_data['id'] == document_id, 'status'] = data.get('status', 'active')
        documents_data.loc[documents_data['id'] == document_id, 'tahun'] = data.get('tahun')
        
        # Save to storage
        success = storage_service.write_csv(documents_data, 'config/aoi-documents.csv')
        
        if success:
            # Return updated document
            updated_row = documents_data[documents_data['id'] == document_id]
            if not updated_row.empty:
                updated_document = safe_serialize_dict(updated_row.iloc[0].to_dict())
                return jsonify(updated_document), 200
            else:
                return jsonify({'error': 'AOI document not found after update'}), 404
        else:
            return jsonify({'error': 'Failed to update AOI document'}), 500
            
    except Exception as e:
        safe_print(f"Error updating AOI document {document_id}: {e}")
        return jsonify({'error': f'Failed to update AOI document: {str(e)}'}), 500

@app.route('/api/aoiDocuments/<string:document_id>', methods=['DELETE'])
def delete_aoi_document(document_id):
    """Delete an AOI document"""
    try:
        # Read existing AOI documents
        documents_data = storage_service.read_csv('config/aoi-documents.csv')
        if documents_data is None:
            return jsonify({'error': 'No AOI documents found'}), 404
        
        # Remove the AOI document
        documents_data = documents_data[documents_data['id'] != document_id]
        
        # Save to storage
        success = storage_service.write_csv(documents_data, 'config/aoi-documents.csv')
        
        if success:
            return jsonify({'message': f'AOI document {document_id} deleted successfully'}), 200
        else:
            return jsonify({'error': 'Failed to delete AOI document'}), 500
            
    except Exception as e:
        safe_print(f"Error deleting AOI document {document_id}: {e}")
        return jsonify({'error': f'Failed to delete AOI document: {str(e)}'}), 500

@app.route('/api/upload-aoi-file', methods=['POST'])
def upload_aoi_file():
    """
    Upload an AOI document file directly to storage.
    This endpoint handles file uploads for Area of Improvement documents.
    """
    try:
        safe_print(f"🔧 DEBUG: AOI file upload request received")
        safe_print(f"🔧 DEBUG: Request files: {list(request.files.keys())}")
        safe_print(f"🔧 DEBUG: Request form: {dict(request.form)}")
        
        # Get the uploaded file
        if 'file' not in request.files:
            return jsonify({'error': 'No file provided'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        safe_print(f"🔧 DEBUG: File received: {file.filename}")
        
        # Get form data
        aoi_recommendation_id = request.form.get('aoiRecommendationId')
        aoi_jenis = request.form.get('aoiJenis', 'REKOMENDASI')
        aoi_urutan = request.form.get('aoiUrutan', '1')
        year = request.form.get('year')
        user_direktorat = request.form.get('userDirektorat', '')
        user_subdirektorat = request.form.get('userSubdirektorat', '')
        user_divisi = request.form.get('userDivisi', '')
        user_id = request.form.get('userId', '')
        
        # Validate required parameters
        if not aoi_recommendation_id or not year:
            return jsonify({'error': 'AOI recommendation ID and year are required'}), 400
        
        try:
            year_int = int(year)
            recommendation_id_int = int(aoi_recommendation_id)
            urutan_int = int(aoi_urutan)
        except ValueError:
            return jsonify({'error': 'Invalid year, recommendation ID, or urutan format'}), 400
        
        # Determine PIC name (use provided or fallback)
        pic_name = user_divisi or user_subdirektorat or user_direktorat or 'Unknown_Division'
        # Replace spaces with underscores for file path
        pic_name_clean = secure_filename(pic_name.replace(' ', '_'))
        
        # Create file path: aoi-documents/{year}/{pic}/{recommendation_id}/{filename}
        filename = secure_filename(file.filename)
        file_path = f"aoi-documents/{year_int}/{pic_name_clean}/{recommendation_id_int}/{filename}"
        
        safe_print(f"🔧 DEBUG: Uploading to path: {file_path}")
        
        # Clear existing files in the recommendation directory first (local filesystem)
        try:
            directory_path = f"aoi-documents/{year_int}/{pic_name_clean}/{recommendation_id_int}"
            safe_print(f"🔧 DEBUG: Clearing directory: {directory_path}")

            # Use local filesystem
            local_dir = Path(__file__).parent.parent / 'data' / directory_path
            if local_dir.exists() and local_dir.is_dir():
                import shutil
                safe_print(f"🔧 DEBUG: Removing existing directory: {local_dir}")
                shutil.rmtree(local_dir)
        except Exception as e:
            safe_print(f"Error clearing directory: {e}")

        # Upload file to local storage
        file_content = file.read()
        local_file_path = Path(__file__).parent.parent / 'data' / file_path

        # Create directory structure
        local_file_path.parent.mkdir(parents=True, exist_ok=True)

        # Save file
        with open(local_file_path, 'wb') as f:
            f.write(file_content)

        safe_print(f"🔧 DEBUG: File uploaded successfully to: {file_path}")
        
        # Create AOI document record
        document_id = f"aoi_{generate_unique_id()}"
        aoi_document_data = {
            'id': document_id,
            'fileName': filename,
            'fileSize': len(file_content),
            'uploadDate': datetime.now().isoformat(),
            'aoiRecommendationId': recommendation_id_int,
            'aoiJenis': aoi_jenis,
            'aoiUrutan': urutan_int,
            'userId': user_id,
            'userDirektorat': user_direktorat,
            'userSubdirektorat': user_subdirektorat,
            'userDivisi': user_divisi,
            'fileType': file.content_type or 'application/octet-stream',
            'status': 'active',
            'tahun': year_int,
            'filePath': file_path
        }
        
        # Read existing AOI documents
        existing_data = storage_service.read_csv('config/aoi-documents.csv')
        if existing_data is not None:
            # Remove existing documents for the same recommendation
            existing_data = existing_data[existing_data['aoiRecommendationId'] != recommendation_id_int]
            documents_df = existing_data
        else:
            documents_df = pd.DataFrame()
        
        # Add new AOI document
        new_document_df = pd.DataFrame([aoi_document_data])
        updated_df = pd.concat([documents_df, new_document_df], ignore_index=True)
        
        # Save to storage
        success = storage_service.write_csv(updated_df, 'config/aoi-documents.csv')
        
        if success:
            safe_print(f"🔧 DEBUG: AOI document record saved successfully")
            return jsonify({
                'message': 'AOI file uploaded successfully',
                'documentId': document_id,
                'filePath': file_path,
                'document': aoi_document_data
            }), 201
        else:
            return jsonify({'error': 'Failed to save AOI document record'}), 500
            
    except Exception as e:
        safe_print(f"🔧 DEBUG: Exception in upload_aoi_file: {e}")
        import traceback
        safe_print(f"🔧 DEBUG: Full traceback: {traceback.format_exc()}")
        return jsonify({'error': f'Failed to upload AOI file: {str(e)}'}), 500

@app.route('/api/users', methods=['GET'])
def get_users():
    """Get all users from storage, optionally filtered by year"""
    try:
        # Get year parameter from query string
        year = request.args.get('year', type=int)

        csv_data = storage_service.read_csv('config/users.csv')
        if csv_data is not None:
            # Replace NaN values with a safe placeholder to prevent empty password login
            csv_data = csv_data.fillna({'password': '[NO_PASSWORD_SET]'}).fillna('')

            # Filter by year if provided
            if year is not None and 'tahun' in csv_data.columns:
                # Filter users that belong to the specified year
                # Also include users without year (tahun == '' or NaN) for backwards compatibility
                csv_data = csv_data[
                    (csv_data['tahun'] == year) |
                    (csv_data['tahun'] == '') |
                    (csv_data['tahun'].isna())
                ]
                safe_print(f"📋 GET /api/users?year={year} - Filtered to {len(csv_data)} users")

            # Ensure WhatsApp field is treated as string, not float
            if 'whatsapp' in csv_data.columns:
                csv_data['whatsapp'] = csv_data['whatsapp'].astype(str).replace('nan', '').str.replace(r'\.0$', '', regex=True)

            users = csv_data.to_dict(orient='records')
            return jsonify(users), 200
        return jsonify([]), 200
    except Exception as e:
        safe_print(f"Error getting users: {e}")
        return jsonify({'error': f'Failed to get users: {str(e)}'}), 500

@app.route('/api/login', methods=['POST'])
def login_user():
    """Authenticate user with email and password"""
    try:
        data = request.get_json()
        email = data.get('email', '').strip()
        password = data.get('password', '').strip()

        if not email or not password:
            return jsonify({'error': 'Email and password are required'}), 400

        # Read users from CSV
        csv_data = storage_service.read_csv('config/users.csv')
        if csv_data is None or csv_data.empty:
            return jsonify({'error': 'No users found'}), 404

        # Find user by email and password
        user_match = csv_data[
            (csv_data['email'].str.lower() == email.lower()) &
            (csv_data['password'] == password) &
            (csv_data['status'] == 'active')
        ]

        if user_match.empty:
            return jsonify({'error': 'Invalid email or password'}), 401

        # Get user data
        user_data = user_match.iloc[0].to_dict()

        # Remove password from response
        if 'password' in user_data:
            del user_data['password']

        # Clean up NaN values and convert to appropriate types
        for key, value in user_data.items():
            if pd.isna(value) or str(value).lower() == 'nan':
                user_data[key] = ''
            else:
                user_data[key] = str(value).replace('.0', '') if isinstance(value, (int, float)) else str(value)

        # Convert id to string and add required fields
        user_data['id'] = str(user_data['id'])
        user_data['createdAt'] = user_data.get('created_at', '')

        return jsonify(user_data), 200

    except Exception as e:
        safe_print(f"Error during login: {e}")
        return jsonify({'error': f'Login failed: {str(e)}'}), 500

@app.route('/api/login-db', methods=['POST'])
def login_user_db():
    """Authenticate user with email and password from CSV file (primary) or SQLite database"""
    from database import get_db_connection
    import bcrypt
    try:
        data = request.get_json()
        email = data.get('email', '').strip()
        password = data.get('password', '').strip()

        safe_print(f"🔐 DEBUG: Login attempt - email: '{email}', password length: {len(password)}")

        if not email or not password:
            safe_print(f"❌ DEBUG: Missing credentials - email: {bool(email)}, password: {bool(password)}")
            return jsonify({'error': 'Email and password are required'}), 400

        # 1. Try CSV file first (primary storage)
        csv_data = storage_service.read_csv('config/users.csv')
        safe_print(f"📋 DEBUG: CSV loaded - rows: {len(csv_data) if csv_data is not None else 'None'}")
        if csv_data is not None and not csv_data.empty:
            # Find user by email in CSV
            user_row = csv_data[csv_data['email'] == email]
            safe_print(f"🔍 DEBUG: User search result - found: {not user_row.empty}")

            if not user_row.empty:
                user = user_row.iloc[0]
                stored_password = str(user.get('password', ''))
                safe_print(f"🔑 DEBUG: Stored password type: {'bcrypt' if stored_password.startswith('$2') else 'plain'}, length: {len(stored_password)}")

                # Check if password is bcrypt hashed or plain text
                password_valid = False
                if stored_password.startswith('$2b$') or stored_password.startswith('$2a$'):
                    # Bcrypt hashed password
                    try:
                        password_valid = bcrypt.checkpw(password.encode('utf-8'), stored_password.encode('utf-8'))
                        safe_print(f"✓ DEBUG: Bcrypt validation result: {password_valid}")
                    except Exception as bcrypt_err:
                        password_valid = False
                        safe_print(f"❌ DEBUG: Bcrypt error: {bcrypt_err}")
                else:
                    # Plain text password (for backward compatibility)
                    password_valid = (password == stored_password)
                    safe_print(f"✓ DEBUG: Plain text validation result: {password_valid}")

                if password_valid:
                    # Prepare user data response (handle NaN values)
                    import math
                    def safe_str(val):
                        if pd.isna(val) or (isinstance(val, float) and math.isnan(val)):
                            return ''
                        return str(val)

                    user_data = {
                        'id': safe_str(user.get('id', '')),
                        'email': safe_str(user.get('email', '')),
                        'role': safe_str(user.get('role', 'user')),
                        'name': safe_str(user.get('name', '')),
                        'direktorat': safe_str(user.get('direktorat', '')),
                        'subdirektorat': safe_str(user.get('subdirektorat', '')),
                        'divisi': safe_str(user.get('divisi', '')),
                        'is_active': True,
                        'createdAt': safe_str(user.get('created_at', ''))
                    }
                    safe_print(f"✅ User logged in from CSV: {email} (Role: {user_data['role']})")
                    return jsonify(user_data), 200

        # 2. If not found in CSV, try SQLite database (fallback)
        try:
            with get_db_connection() as conn:
                cursor = conn.cursor()

                # Find user by email
                cursor.execute("""
                    SELECT id, email, password_hash, role, name, direktorat, subdirektorat, divisi, is_active
                    FROM users
                    WHERE email = ? AND is_active = 1
                """, (email,))

                user = cursor.fetchone()

                if user:
                    user_id, user_email, password_hash, role, name, direktorat, subdirektorat, divisi, is_active = user

                    # Verify password with bcrypt
                    if bcrypt.checkpw(password.encode('utf-8'), password_hash.encode('utf-8')):
                        # Prepare user data response
                        user_data = {
                            'id': str(user_id),
                            'email': user_email,
                            'role': role,
                            'name': name or '',
                            'direktorat': direktorat or '',
                            'subdirektorat': subdirektorat or '',
                            'divisi': divisi or '',
                            'is_active': bool(is_active),
                            'createdAt': ''
                        }
                        safe_print(f"✅ User logged in from SQLite: {user_email} (Role: {role})")
                        return jsonify(user_data), 200
        except Exception as db_error:
            safe_print(f"⚠️ SQLite login failed: {db_error}")

        # 3. User not found or password invalid
        safe_print(f"❌ Login failed for: {email}")
        return jsonify({'error': 'Invalid email or password'}), 401

    except Exception as e:
        safe_print(f"❌ Error during login: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'Login failed: {str(e)}'}), 500

@app.route('/api/users/<string:user_id>', methods=['GET'])
def get_user_by_id(user_id):
    """Get a specific user by ID from storage"""
    try:
        csv_data = storage_service.read_csv('config/users.csv')
        if csv_data is None:
            return jsonify({'error': 'No users found'}), 404
        
        # Find user by ID (handle both string and int IDs)
        user_row = csv_data[csv_data['id'].astype(str) == str(user_id)]
        
        if user_row.empty:
            return jsonify({'error': 'User not found'}), 404
        
        # Return the user data
        user_data = safe_serialize_dict(user_row.iloc[0].to_dict())
        return jsonify(user_data), 200
        
    except Exception as e:
        safe_print(f"Error getting user by ID {user_id}: {e}")
        return jsonify({'error': f'Failed to get user: {str(e)}'}), 500

@app.route('/api/users', methods=['POST'])
def create_user():
    """Create a new user and save to CSV file (primary storage)"""
    try:
        data = request.get_json()
        year = data.get('tahun') or data.get('year')  # Support both field names

        # Validate required fields
        if not data.get('email'):
            return jsonify({'error': 'Email is required'}), 400
        if not data.get('password'):
            return jsonify({'error': 'Password is required'}), 400

        # Read existing users from CSV
        csv_data = storage_service.read_csv('config/users.csv')
        if csv_data is None:
            csv_data = pd.DataFrame()

        # Check if email already exists for this year
        if not csv_data.empty and 'email' in csv_data.columns:
            existing = csv_data[csv_data['email'].str.lower() == data.get('email', '').lower()]
            if not existing.empty:
                # If year is specified, check if user exists in this year
                if year and 'tahun' in csv_data.columns:
                    existing_in_year = existing[existing['tahun'] == year]
                    if not existing_in_year.empty:
                        return jsonify({'error': 'User with this email already exists for this year'}), 400
                else:
                    return jsonify({'error': 'User with this email already exists'}), 400

        # Generate user ID
        import time
        user_id = int(time.time() * 1000)

        # Create new user record
        new_user = {
            'id': user_id,
            'name': data.get('name', ''),
            'email': data.get('email'),
            'password': data.get('password'),  # Store plain password in CSV
            'role': data.get('role', 'user'),
            'direktorat': data.get('direktorat', ''),
            'subdirektorat': data.get('subdirektorat', ''),
            'divisi': data.get('divisi', ''),
            'status': 'active',
            'tahun': year if year else '',  # Save year!
            'created_at': datetime.now().isoformat(),
            'is_active': 1,
            'whatsapp': data.get('whatsapp', '')
        }

        # Add to DataFrame
        new_row = pd.DataFrame([new_user])
        csv_data = pd.concat([csv_data, new_row], ignore_index=True)

        # Save to CSV
        success = storage_service.write_csv(csv_data, 'config/users.csv')

        if success:
            safe_print(f"✅ Created user in CSV: {new_user['email']} (ID: {user_id}, Year: {year or 'N/A'})")

            # Return user data without password
            response_data = new_user.copy()
            response_data.pop('password', None)
            return jsonify(response_data), 201
        else:
            return jsonify({'error': 'Failed to save user'}), 500

    except Exception as e:
        safe_print(f"Error creating user: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'Failed to create user: {str(e)}'}), 500

@app.route('/api/users/<int:user_id>', methods=['DELETE'])
def delete_user(user_id):
    """Delete a user from CSV file and SQLite database"""
    from database import get_db_connection
    try:
        # 1. Delete from CSV file (primary storage)
        csv_data = storage_service.read_csv('config/users.csv')
        if csv_data is None or csv_data.empty:
            return jsonify({'error': 'No users data found'}), 404

        # Find user in CSV
        user_exists = csv_data[csv_data['id'] == user_id]
        if user_exists.empty:
            return jsonify({'error': 'User not found'}), 404

        user_email = user_exists.iloc[0]['email']

        # Remove user from CSV
        csv_data = csv_data[csv_data['id'] != user_id]
        storage_service.write_csv(csv_data, 'config/users.csv')
        safe_print(f"✅ Deleted user from CSV: {user_email} (ID: {user_id})")

        # 2. Also delete from SQLite database if exists
        try:
            with get_db_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT id FROM users WHERE id = ?", (user_id,))
                if cursor.fetchone():
                    cursor.execute("UPDATE users SET is_active = 0, updated_at = CURRENT_TIMESTAMP WHERE id = ?", (user_id,))
                    safe_print(f"✅ Soft-deleted user from SQLite: {user_email} (ID: {user_id})")
        except Exception as db_error:
            safe_print(f"Warning: Could not delete from SQLite: {db_error}")

        return jsonify({'message': 'User deleted successfully'}), 200

    except Exception as e:
        safe_print(f"Error deleting user: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'Failed to delete user: {str(e)}'}), 500

@app.route('/api/users/<int:user_id>', methods=['PUT'])
def update_user(user_id):
    """Update user in CSV file and SQLite database"""
    from database import get_db_connection
    try:
        data = request.get_json()

        # 1. Update CSV file (primary storage)
        csv_data = storage_service.read_csv('config/users.csv')
        if csv_data is None or csv_data.empty:
            return jsonify({'error': 'No users data found'}), 404

        # Find user in CSV
        user_mask = csv_data['id'] == user_id
        if not user_mask.any():
            return jsonify({'error': 'User not found'}), 404

        user_index = csv_data[user_mask].index[0]

        # Update fields in CSV
        if 'name' in data:
            csv_data.at[user_index, 'name'] = data['name']
        if 'email' in data:
            csv_data.at[user_index, 'email'] = data['email']
        if 'password' in data:
            # Hash password for CSV storage
            try:
                import bcrypt
                password_hash = bcrypt.hashpw(data['password'].encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
                csv_data.at[user_index, 'password'] = password_hash
            except:
                csv_data.at[user_index, 'password'] = data['password']
        if 'role' in data:
            csv_data.at[user_index, 'role'] = data['role']
        if 'direktorat' in data:
            csv_data.at[user_index, 'direktorat'] = data['direktorat']
        if 'subdirektorat' in data:
            csv_data.at[user_index, 'subdirektorat'] = data['subdirektorat']
        if 'divisi' in data:
            csv_data.at[user_index, 'divisi'] = data['divisi']
        if 'tahun' in data or 'year' in data:
            csv_data.at[user_index, 'tahun'] = data.get('tahun') or data.get('year', '')
        if 'whatsapp' in data:
            csv_data.at[user_index, 'whatsapp'] = data['whatsapp']
        if 'telegram' in data:
            csv_data.at[user_index, 'telegram'] = data['telegram']

        # Save updated CSV
        storage_service.write_csv(csv_data, 'config/users.csv')

        # Get updated user data from CSV
        updated_user_row = csv_data.iloc[user_index]
        updated_user = {
            'id': int(updated_user_row['id']),
            'name': updated_user_row.get('name', ''),
            'email': updated_user_row.get('email', ''),
            'role': updated_user_row.get('role', 'user'),
            'direktorat': updated_user_row.get('direktorat', ''),
            'subdirektorat': updated_user_row.get('subdirektorat', ''),
            'divisi': updated_user_row.get('divisi', ''),
            'tahun': updated_user_row.get('tahun', '')
        }

        safe_print(f"✅ Updated user in CSV: {updated_user['email']} (ID: {user_id})")

        # 2. Also update SQLite database if user exists there
        try:
            with get_db_connection() as conn:
                cursor = conn.cursor()

                # Check if user exists in SQLite
                cursor.execute("SELECT id FROM users WHERE id = ?", (user_id,))
                if cursor.fetchone():
                    # Build dynamic UPDATE query
                    update_fields = []
                    update_values = []

                    if 'name' in data:
                        update_fields.append('name = ?')
                        update_values.append(data['name'])
                    if 'email' in data:
                        update_fields.append('email = ?')
                        update_values.append(data['email'])
                    if 'password' in data:
                        # Hash password for SQLite
                        import bcrypt
                        password_hash = bcrypt.hashpw(data['password'].encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
                        update_fields.append('password_hash = ?')
                        update_values.append(password_hash)
                    if 'role' in data:
                        update_fields.append('role = ?')
                        update_values.append(data['role'])
                    if 'direktorat' in data:
                        update_fields.append('direktorat = ?')
                        update_values.append(data['direktorat'])
                    if 'subdirektorat' in data:
                        update_fields.append('subdirektorat = ?')
                        update_values.append(data['subdirektorat'])
                    if 'divisi' in data:
                        update_fields.append('divisi = ?')
                        update_values.append(data['divisi'])
                    if 'whatsapp' in data:
                        update_fields.append('whatsapp = ?')
                        update_values.append(data['whatsapp'])
                    if 'telegram' in data:
                        update_fields.append('telegram = ?')
                        update_values.append(data['telegram'])

                    # Add updated_at timestamp
                    update_fields.append('updated_at = CURRENT_TIMESTAMP')

                    if update_fields:
                        update_query = f"UPDATE users SET {', '.join(update_fields)} WHERE id = ?"
                        update_values.append(user_id)
                        cursor.execute(update_query, update_values)
                        safe_print(f"✅ Updated user in SQLite: {updated_user['email']} (ID: {user_id})")
        except Exception as db_error:
            safe_print(f"⚠️ Warning: Could not update in SQLite (user may not exist in DB): {db_error}")

        return jsonify(updated_user), 200

    except Exception as e:
        safe_print(f"❌ Error updating user: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'Failed to update user: {str(e)}'}), 500

@app.route('/api/upload-gcg-file', methods=['POST'])
def upload_gcg_file():
    """
    Upload a GCG document file directly to storage.
    This endpoint is specifically for the monitoring upload feature.
    """
    try:
        safe_print(f"🔧 DEBUG: GCG file upload request received")
        safe_print(f"🔧 DEBUG: Request files: {list(request.files.keys())}")
        safe_print(f"🔧 DEBUG: Request form: {dict(request.form)}")
        
        # Check if file is present
        if 'file' not in request.files:
            safe_print(f"🔧 DEBUG: No file in request")
            return jsonify({'error': 'No file provided'}), 400
        
        file = request.files['file']
        safe_print(f"🔧 DEBUG: File received: {file.filename}")
        
        if file.filename == '':
            safe_print(f"🔧 DEBUG: Empty filename")
            return jsonify({'error': 'No file selected'}), 400
        
        # Get metadata from form
        year = request.form.get('year')
        checklist_id = request.form.get('checklistId')
        checklist_description = request.form.get('checklistDescription', '')
        aspect = request.form.get('aspect', '')
        subdirektorat = request.form.get('subdirektorat', '')
        catatan = request.form.get('catatan', '')  # Catatan from user
        row_number = request.form.get('rowNumber')  # New: row number for document organization
        
        # Validate required fields
        if not year:
            return jsonify({'error': 'Year is required'}), 400
        if not checklist_id:
            return jsonify({'error': 'Checklist ID is required'}), 400
        
        try:
            year_int = int(year)
            checklist_id_int = int(checklist_id)
        except ValueError:
            return jsonify({'error': 'Invalid year or checklist ID format'}), 400
        
        # Generate file ID for record tracking
        file_id = str(uuid.uuid4())
        
        # Clean subdirektorat name for use in file path
        pic_name = secure_filename(subdirektorat) if subdirektorat else 'UNKNOWN_PIC'
        
        # Fixed file structure: gcg-documents/{year}/{PIC}/{checklist_id}/{filename}
        file_path = f"gcg-documents/{year_int}/{pic_name}/{checklist_id_int}/{secure_filename(file.filename)}"
        
        # Upload file to storage
        file_data = file.read()
        
        # Determine content type
        file_extension = file.filename.rsplit('.', 1)[1].lower() if '.' in file.filename else 'bin'
        content_type_map = {
            'pdf': 'application/pdf',
            'doc': 'application/msword',
            'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'xls': 'application/vnd.ms-excel',
            'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'ppt': 'application/vnd.ms-powerpoint',
            'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        }
        content_type = content_type_map.get(file_extension, 'application/octet-stream')
        
        # Upload to local storage using organized directory structure
        try:
            safe_print(f"🔧 DEBUG: Uploading to local storage path: {file_path}")

            # Create local file path
            local_file_path = Path(__file__).parent.parent / 'data' / file_path
            safe_print(f"🔧 DEBUG: Full local path: {local_file_path}")

            # First, delete ALL existing files in the directory to ensure clean overwrite
            try:
                # Clear the directory (keep the directory structure clean)
                directory_path = local_file_path.parent
                safe_print(f"🔧 DEBUG: Clearing directory: {directory_path}")

                if directory_path.exists():
                    # Remove all files in the directory but keep the directory structure
                    for existing_file in directory_path.glob('*'):
                        if existing_file.is_file():
                            safe_print(f"🔧 DEBUG: Removing existing file: {existing_file}")
                            existing_file.unlink()
                else:
                    safe_print(f"🔧 DEBUG: Directory doesn't exist, will be created")

            except Exception as e:
                safe_print(f"🔧 DEBUG: Error clearing directory (continuing anyway): {e}")

            # Create directory structure if it doesn't exist
            local_file_path.parent.mkdir(parents=True, exist_ok=True)

            # Save the file to local storage
            with open(local_file_path, 'wb') as f:
                f.write(file_data)

            safe_print(f"🔧 DEBUG: File saved successfully to local storage: {local_file_path}")

        except Exception as upload_error:
            safe_print(f"🔧 DEBUG: Local upload exception: {upload_error}")
            return jsonify({'error': f'Failed to save file to local storage: {str(upload_error)}'}), 500
        
        # Get user information from form (if provided)
        uploaded_by = request.form.get('uploadedBy', 'Unknown User')
        user_role = request.form.get('userRole', 'user')  # Default to 'user' role
        user_direktorat = request.form.get('userDirektorat', 'Unknown')
        user_subdirektorat = request.form.get('userSubdirektorat', 'Unknown')
        user_divisi = request.form.get('userDivisi', 'Unknown')
        user_whatsapp = request.form.get('userWhatsApp', '')
        user_email = request.form.get('userEmail', '')
        
        # Create file record
        file_record = {
            'id': file_id,
            'fileName': file.filename,
            'fileSize': len(file_data),
            'uploadDate': datetime.now().isoformat(),
            'year': year_int,
            'checklistId': int(float(checklist_id)) if checklist_id else None,
            'checklistDescription': checklist_description,
            'aspect': aspect,
            'subdirektorat': subdirektorat,
            'status': 'uploaded',
            'localFilePath': file_path,  # Keep same path structure for compatibility
            'uploadedBy': uploaded_by,
            'userRole': user_role,
            'userDirektorat': user_direktorat,
            'userSubdirektorat': user_subdirektorat,
            'userDivisi': user_divisi,
            'userWhatsApp': user_whatsapp,
            'userEmail': user_email,
            'catatan': catatan
        }
        
        # Save to SQLite database (primary storage)
        from database import get_db_connection
        try:
            with get_db_connection() as conn:
                cursor = conn.cursor()

                # Remove existing record for this checklistId and year (for re-upload scenario)
                cursor.execute("""
                    DELETE FROM uploaded_files
                    WHERE checklist_id = ? AND year = ?
                """, (checklist_id_int, year_int))
                deleted_count = cursor.rowcount
                if deleted_count > 0:
                    safe_print(f"🔧 DEBUG: Deleted {deleted_count} existing record(s) for re-upload")

                # Insert new record
                cursor.execute("""
                    INSERT INTO uploaded_files (
                        id, file_name, file_size, upload_date, year,
                        checklist_id, checklist_description, aspect, status, file_path
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    file_id,
                    file.filename,
                    len(file_data),
                    datetime.now().isoformat(),
                    year_int,
                    checklist_id_int,
                    checklist_description,
                    aspect,
                    'uploaded',
                    file_path
                ))

                conn.commit()
                safe_print(f"🔧 DEBUG: File record saved to database successfully")

        except Exception as db_error:
            safe_print(f"🔧 DEBUG: Error saving to database: {db_error}")
            import traceback
            safe_print(f"🔧 DEBUG: Database error traceback: {traceback.format_exc()}")
            return jsonify({'error': f'File uploaded but failed to save database record: {str(db_error)}'}), 500

        # Also save to Excel file for backward compatibility (legacy support)
        try:
            files_data = storage_service.read_excel('uploaded-files.xlsx')
            if files_data is None:
                files_data = pd.DataFrame()

            # Remove existing record
            if not files_data.empty:
                files_data = files_data[~((files_data['checklistId'] == checklist_id_int) & (files_data['year'] == year_int))]

            new_row = pd.DataFrame([file_record])
            files_data = pd.concat([files_data, new_row], ignore_index=True)
            storage_service.write_excel(files_data, 'uploaded-files.xlsx')
            safe_print(f"🔧 DEBUG: Also saved to Excel for backward compatibility")
        except Exception as excel_error:
            # Non-critical - Excel is just backup
            safe_print(f"🔧 WARNING: Could not save to Excel (non-critical): {excel_error}")

        return jsonify({
            'success': True,
            'file': file_record,
            'message': 'File uploaded successfully to local storage'
        }), 201
        
    except Exception as e:
        safe_print(f"🔧 DEBUG: Exception in upload_gcg_file: {e}")
        import traceback
        safe_print(f"🔧 DEBUG: Full traceback: {traceback.format_exc()}")
        return jsonify({'error': f'Failed to upload GCG file: {str(e)}'}), 500

@app.route('/api/upload-random-document', methods=['POST'])
def upload_random_document():
    """
    Upload random/unstructured document to Dokumen Lainnya folder.
    These documents don't have checklist assignments or organizational structure.
    """
    safe_print(f"📤 DEBUG: Random document upload request received")
    safe_print(f"📋 DEBUG: Content-Length: {request.content_length}")
    safe_print(f"📋 DEBUG: Content-Type: {request.content_type}")

    try:
        # File size check removed - now handled by Flask MAX_CONTENT_LENGTH (500MB)

        # Check if file is present
        if 'file' not in request.files:
            safe_print("❌ DEBUG: No file in request.files")
            return jsonify({'error': 'No file provided'}), 400

        file = request.files['file']
        safe_print(f"📦 DEBUG: File received: {file.filename}")

        if file.filename == '':
            safe_print("❌ DEBUG: Empty filename")
            return jsonify({'error': 'No file selected'}), 400

        # Log file size for debugging
        try:
            file.seek(0, 2)  # Seek to end
            file_size = file.tell()
            file.seek(0)  # Reset to beginning
            safe_print(f"📦 DEBUG: File size: {file_size / (1024 * 1024):.2f}MB")
            # File size limit removed - now handled by Flask MAX_CONTENT_LENGTH (500MB)
        except Exception as seek_error:
            safe_print(f"⚠️ DEBUG: Could not check file size: {seek_error}")
            # Continue anyway

        # Get metadata from form
        year = request.form.get('year')
        category = request.form.get('category', 'dokumen_lainnya')
        uploaded_by = request.form.get('uploadedBy', 'Unknown User')
        folder_path = request.form.get('folderPath', '')  # Preserve folder structure

        # Validate required fields
        if not year:
            return jsonify({'error': 'Year is required'}), 400

        try:
            year_int = int(year)
        except ValueError:
            return jsonify({'error': 'Invalid year format'}), 400

        # Generate file ID for record tracking
        file_id = str(uuid.uuid4())

        # File structure with folder preservation
        safe_filename_str = secure_filename(file.filename)

        if folder_path:
            # Preserve folder structure from upload
            # Remove first component (folder name itself) and keep subdirectories
            path_parts = folder_path.split('/')
            if len(path_parts) > 1:
                # Keep subdirectory structure
                folder_structure = '/'.join(path_parts[:-1])  # Remove filename
                safe_folder = secure_filename(folder_structure.replace('/', '_'))
                file_path = f"gcg-documents/{year_int}/Dokumen_Lainnya/{safe_folder}/{safe_filename_str}"
            else:
                file_path = f"gcg-documents/{year_int}/Dokumen_Lainnya/{safe_filename_str}"
        else:
            # Single file upload - add timestamp to prevent conflicts
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename_parts = safe_filename_str.rsplit('.', 1)
            if len(filename_parts) == 2:
                unique_filename = f"{filename_parts[0]}_{timestamp}.{filename_parts[1]}"
            else:
                unique_filename = f"{safe_filename_str}_{timestamp}"
            file_path = f"gcg-documents/{year_int}/Dokumen_Lainnya/{unique_filename}"

        # Upload file to local storage
        file_data = file.read()

        try:
            local_file_path = Path(__file__).parent.parent / 'data' / file_path
            safe_print(f"📤 DEBUG: Saving to: {local_file_path}")

            # Create directory structure
            local_file_path.parent.mkdir(parents=True, exist_ok=True)

            # Save the file
            with open(local_file_path, 'wb') as f:
                f.write(file_data)

            safe_print(f"✅ DEBUG: File saved successfully: {local_file_path}")

        except Exception as upload_error:
            safe_print(f"❌ DEBUG: Upload error: {upload_error}")
            return jsonify({'error': f'Failed to save file: {str(upload_error)}'}), 500

        # Create file record
        file_record = {
            'id': file_id,
            'fileName': file.filename,
            'fileSize': len(file_data),
            'uploadDate': datetime.now().isoformat(),
            'year': year_int,
            'checklistId': None,  # No checklist association
            'checklistDescription': 'Dokumen Lainnya - Arsip',
            'aspect': 'DOKUMEN_LAINNYA',
            'subdirektorat': 'Dokumen_Lainnya',
            'status': 'uploaded',
            'localFilePath': file_path,
            'uploadedBy': uploaded_by,
            'userRole': 'admin',
            'catatan': f'Uploaded to Dokumen Lainnya folder on {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}'
        }

        # Add to uploaded files database
        try:
            files_data = storage_service.read_excel('uploaded-files.xlsx')
            if files_data is None:
                files_data = pd.DataFrame()
        except Exception:
            files_data = pd.DataFrame()

        new_row = pd.DataFrame([file_record])
        files_data = pd.concat([files_data, new_row], ignore_index=True)

        # Save to storage
        try:
            success = storage_service.write_excel(files_data, 'uploaded-files.xlsx')

            if success:
                return jsonify({
                    'success': True,
                    'file': file_record,
                    'message': 'Random document uploaded successfully'
                }), 201
            else:
                return jsonify({'error': 'File uploaded but failed to save record'}), 500
        except Exception as save_error:
            safe_print(f"❌ DEBUG: Save error: {save_error}")
            return jsonify({'error': f'Failed to save record: {str(save_error)}'}), 500

    except Exception as e:
        safe_print(f"❌ DEBUG: Exception in upload_random_document: {e}")
        import traceback
        safe_print(f"❌ DEBUG: Traceback: {traceback.format_exc()}")
        return jsonify({'error': f'Failed to upload random document: {str(e)}'}), 500

@app.route('/api/random-documents/<int:year>', methods=['GET'])
def get_random_documents(year):
    """Get all random documents (Dokumen Lainnya) for a specific year"""
    try:
        safe_print(f"📂 DEBUG: Loading random documents for year {year}")

        # Load uploaded files data
        try:
            files_data = storage_service.read_excel('uploaded-files.xlsx')
            if files_data is None or files_data.empty:
                safe_print(f"⚠️ DEBUG: No uploaded files found")
                return jsonify({'documents': []}), 200
        except Exception as e:
            safe_print(f"❌ DEBUG: Error reading uploaded-files.xlsx: {e}")
            return jsonify({'documents': []}), 200

        # Filter for random documents (checklistId is null/empty AND year matches)
        random_docs = files_data[
            (files_data['year'] == year) &
            (files_data['checklistId'].isna() | (files_data['checklistId'] == ''))
        ]

        safe_print(f"✅ DEBUG: Found {len(random_docs)} random documents for year {year}")

        # Convert to list of dictionaries
        documents = []
        for _, row in random_docs.iterrows():
            doc = {
                'id': row.get('id', ''),
                'fileName': row.get('fileName', ''),
                'fileSize': int(row.get('fileSize', 0)) if pd.notna(row.get('fileSize')) else 0,
                'uploadDate': row.get('uploadDate', ''),
                'uploadedBy': row.get('uploadedBy', 'Unknown'),
                'subdirektorat': row.get('subdirektorat', 'Dokumen_Lainnya'),
                'catatan': row.get('catatan', ''),
                'localFilePath': row.get('localFilePath', ''),
                'year': int(row.get('year', year))
            }
            documents.append(doc)

        return jsonify({'documents': documents}), 200

    except Exception as e:
        safe_print(f"❌ DEBUG: Exception in get_random_documents: {e}")
        import traceback
        safe_print(f"❌ DEBUG: Traceback: {traceback.format_exc()}")
        return jsonify({'error': f'Failed to load random documents: {str(e)}'}), 500

@app.route('/api/check-gcg-files', methods=['POST'])
def check_gcg_files():
    """Check if GCG files exist by querying uploaded_files table (fast database lookup)"""
    try:
        data = request.get_json()
        safe_print(f"🔍 DEBUG: check_gcg_files received data: {data}")

        pic_name = data.get('picName')
        checklist_ids = data.get('checklistIds', [])  # List of checklist IDs to check
        year = data.get('year')  # Year is required for new structure
        verify_files = data.get('verifyFiles', False)  # Optional: verify filesystem (slower but accurate)

        safe_print(f"🔍 DEBUG: pic_name={pic_name}, year={year}, checklist_ids={checklist_ids}, verify_files={verify_files}")

        # Support legacy row_numbers parameter for backward compatibility
        if not checklist_ids and data.get('rowNumbers'):
            safe_print(f"🔍 DEBUG: Using legacy rowNumbers: {data.get('rowNumbers')}")
            # Convert row numbers to checklist IDs (year_prefix * 10 + row_number)
            year_prefix = int(str(year)[-2:])
            row_numbers = data.get('rowNumbers', [])
            safe_print(f"🔍 DEBUG: Converting rowNumbers {row_numbers} with year_prefix {year_prefix}")
            checklist_ids = []
            for row_number in row_numbers:
                try:
                    safe_print(f"🔍 DEBUG: Processing row_number: {row_number} (type: {type(row_number)})")
                    checklist_id = year_prefix * 10 + int(row_number)
                    checklist_ids.append(checklist_id)
                except ValueError as ve:
                    safe_print(f"❌ ERROR: Cannot convert row_number {row_number} to int: {ve}")
                    return jsonify({'error': f'Invalid row number: {row_number}'}), 400

        if not checklist_ids or not year:
            return jsonify({'error': 'Year and checklist IDs are required'}), 400

        # Query database for uploaded files - MUCH faster than filesystem scanning
        from database import get_db_connection
        file_statuses = {}

        with get_db_connection() as conn:
            cursor = conn.cursor()

            # Build IN clause for checklist_ids
            placeholders = ','.join('?' * len(checklist_ids))
            query = f"""
                SELECT
                    uf.id,
                    uf.checklist_id,
                    uf.file_name,
                    uf.file_size,
                    uf.upload_date,
                    uf.aspect,
                    uf.checklist_description,
                    uf.status,
                    u.name as uploaded_by,
                    u.subdirektorat,
                    uf.file_path
                FROM uploaded_files uf
                LEFT JOIN users u ON uf.id LIKE '%' || u.email || '%'
                WHERE uf.year = ?
                AND uf.checklist_id IN ({placeholders})
                AND uf.status = 'uploaded'
            """

            # Execute query with year + checklist_ids
            cursor.execute(query, [year] + checklist_ids)
            rows = cursor.fetchall()

            # Convert rows to dict for easy lookup
            uploaded_files_map = {}
            for row in rows:
                checklist_id = row[1]
                uploaded_files_map[checklist_id] = {
                    'id': row[0],
                    'checklistId': row[1],
                    'fileName': row[2],
                    'size': row[3],
                    'uploadDate': row[4],
                    'aspect': row[5] or '',
                    'checklistDescription': row[6] or '',
                    'status': row[7],
                    'uploadedBy': row[8] or 'Unknown',
                    'subdirektorat': row[9] or '',
                    'filePath': row[10] or '',  # Actual stored file path
                }

            # Build response for each checklist_id
            for checklist_id in checklist_ids:
                if checklist_id in uploaded_files_map:
                    file_info = uploaded_files_map[checklist_id]
                    # Use stored file path from database (handles PIC changes correctly)
                    stored_file_path = file_info.get('filePath')

                    if not stored_file_path:
                        # Fallback: construct path for old records without file_path
                        from werkzeug.utils import secure_filename
                        pic_name_clean = secure_filename(pic_name.replace(' ', '_'))
                        stored_file_path = f"gcg-documents/{year}/{pic_name_clean}/{checklist_id}/{file_info['fileName']}"

                    # OPTIONAL: Verify file actually exists on filesystem (only if verifyFiles=true)
                    # This is slower but detects orphaned database records
                    if verify_files:
                        local_file_path = Path(__file__).parent.parent / 'data' / stored_file_path
                        file_actually_exists = local_file_path.exists() and local_file_path.is_file()

                        if not file_actually_exists:
                            # Database record exists but file is missing - orphaned record!
                            safe_print(f"⚠️ WARNING: Orphaned database record for checklist_id {checklist_id} - file missing: {local_file_path}")
                            file_statuses[str(checklist_id)] = {
                                'exists': False,
                                'orphanedRecord': True,  # Flag for cleanup
                                'databaseId': file_info['id']
                            }
                            continue

                    # File exists in database (and filesystem if verified) - return metadata
                    file_statuses[str(checklist_id)] = {
                        'exists': True,
                        'fileName': file_info['fileName'],
                        'path': stored_file_path,
                        'size': file_info['size'],
                        'lastModified': file_info['uploadDate'],
                        'uploadedBy': file_info['uploadedBy'],
                        'subdirektorat': file_info['subdirektorat'],
                        'aspect': file_info['aspect'],
                        'checklistDescription': file_info['checklistDescription'],
                        'checklistId': checklist_id,
                        'catatan': '',  # TODO: Add catatan field to uploaded_files table
                        'id': file_info['id'],
                        'verified': verify_files  # Flag to show if filesystem was checked
                    }
                else:
                    file_statuses[str(checklist_id)] = {'exists': False}
        
        return jsonify({
            'year': year,
            'picName': pic_name,
            'fileStatuses': file_statuses
        }), 200
        
    except Exception as e:
        safe_print(f"Error checking GCG files: {e}")
        return jsonify({'error': f'Failed to check files: {str(e)}'}), 500

@app.route('/api/download-gcg-file', methods=['POST'])
def download_gcg_file():
    """Download GCG file from storage"""
    try:
        # Handle both JSON and form data
        if request.is_json:
            data = request.get_json()
            pic_name = data.get('picName')
            year = data.get('year')
            # Support both rowNumber (legacy) and checklistId (new)
            row_number = data.get('rowNumber')
            checklist_id = data.get('checklistId')
        else:
            # Handle form data
            pic_name = request.form.get('picName')
            year = request.form.get('year')
            row_number = request.form.get('rowNumber')
            checklist_id = request.form.get('checklistId')
            # Convert to int for year and identifiers
            if year:
                year = int(year)
            if row_number:
                row_number = int(row_number)
            if checklist_id:
                checklist_id = int(checklist_id)
        
        # Use checklistId if provided, otherwise fall back to rowNumber
        folder_id = checklist_id if checklist_id else row_number
        
        if not all([pic_name, year, folder_id]):
            return jsonify({'error': 'PIC name, year, and checklist ID (or row number) are required'}), 400
        
        # Clean PIC name
        pic_name_clean = secure_filename(pic_name)
        
        # Use local storage with local storage
        folder_path = f"gcg-documents/{year}/{pic_name_clean}/{folder_id}"

        try:
            # Check local directory for files
            local_folder_path = Path(__file__).parent.parent / 'data' / folder_path

            if not local_folder_path.exists() or not local_folder_path.is_dir():
                return jsonify({'error': 'File directory not found'}), 404

            # Get all files in the directory (excluding hidden files)
            real_files = [f for f in local_folder_path.iterdir() if f.is_file() and not f.name.startswith('.')]

            if not real_files:
                return jsonify({'error': 'No files found in directory'}), 404

            # Get the first (and should be only) real file in the directory
            file_path_obj = real_files[0]
            file_name = file_path_obj.name

            # Read the file from local storage
            try:
                with open(file_path_obj, 'rb') as f:
                    file_response = f.read()
            except Exception as e:
                return jsonify({'error': f'Failed to read file from local storage: {str(e)}'}), 500
            
            # Return the file as a download with proper MIME type detection
            import mimetypes
            
            # Detect proper MIME type from file extension
            mime_type, _ = mimetypes.guess_type(file_name)
            if not mime_type:
                mime_type = 'application/octet-stream'  # Default binary type
            
            response = make_response(file_response)
            response.headers['Content-Disposition'] = f'attachment; filename="{file_name}"'
            response.headers['Content-Type'] = mime_type
            response.headers['Content-Transfer-Encoding'] = 'binary'
            response.headers['Content-Length'] = str(len(file_response))
            response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
            response.headers['Pragma'] = 'no-cache'
            response.headers['Expires'] = '0'
            response.headers['X-Content-Type-Options'] = 'nosniff'
            return response
            
        except Exception as e:
            return jsonify({'error': f'File not found or download failed: {str(e)}'}), 404
        
    except Exception as e:
        safe_print(f"Error downloading GCG file: {e}")
        return jsonify({'error': f'Failed to download file: {str(e)}'}), 500


# ========================
# OLD CONFIGURATION MANAGEMENT ENDPOINTS - DISABLED
# ========================
# These old routes use CSV/storage_service and are now DISABLED.
# The new SQLite-based routes are in api_config_routes.py (registered as blueprints).
# To re-enable these old routes, change @app_OLD.route back to @app.route
# ========================

@app.route('/api/config/aspects', methods=['GET'])
def get_aspects():
    """Get all aspects, optionally filtered by year"""
    try:
        from database import get_db_connection
        year = request.args.get('year')

        # Read aspects from SQLite database (NEW)
        with get_db_connection() as conn:
            cursor = conn.cursor()
            if year:
                year_int = int(year)
                cursor.execute("""
                    SELECT id, nama, deskripsi, tahun, urutan, is_active, created_at
                    FROM aspek_master
                    WHERE tahun = ? AND is_active = 1
                    ORDER BY urutan, nama
                """, (year_int,))
            else:
                cursor.execute("""
                    SELECT id, nama, deskripsi, tahun, urutan, is_active, created_at
                    FROM aspek_master
                    WHERE is_active = 1
                    ORDER BY tahun, urutan, nama
                """)
            rows = cursor.fetchall()
            aspects_list = [dict(row) for row in rows]

        return jsonify({'aspects': aspects_list}), 200

    except Exception as e:
        safe_print(f"Error getting aspects: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'aspects': []}), 200

@app.route('/api/config/aspects', methods=['POST'])
def add_aspect():
    """Add a new aspect for a specific year"""
    try:
        data = request.get_json()
        
        if not data.get('nama') or not data.get('tahun'):
            return jsonify({'error': 'Name and year are required'}), 400
            
        # Read existing aspects
        try:
            aspects_data = storage_service.read_csv('config/aspects.csv')
            if aspects_data is None:
                aspects_data = pd.DataFrame()
        except:
            aspects_data = pd.DataFrame()
            
        # Create new aspect record
        new_aspect = {
            'id': int(time.time() * 1000),  # Use timestamp as ID
            'nama': data['nama'],
            'tahun': int(data['tahun']),
            'created_at': datetime.now().isoformat()
        }
        
        # Add to DataFrame
        new_row = pd.DataFrame([new_aspect])
        aspects_data = pd.concat([aspects_data, new_row], ignore_index=True)
        
        # Save to storage (CSV for easier reading)
        success = storage_service.write_csv(aspects_data, 'config/aspects.csv')
        
        if success:
            return jsonify({'success': True, 'aspect': new_aspect}), 201
        else:
            return jsonify({'error': 'Failed to save aspect'}), 500
            
    except Exception as e:
        safe_print(f"Error adding aspect: {e}")
        return jsonify({'error': f'Failed to add aspect: {str(e)}'}), 500

@app.route('/api/config/aspects/<int:aspect_id>', methods=['PUT'])
def update_aspect(aspect_id):
    """Update an existing aspect"""
    try:
        data = request.get_json()
        
        if not data.get('nama'):
            return jsonify({'error': 'Name is required'}), 400
            
        # Read existing aspects
        aspects_data = storage_service.read_csv('config/aspects.csv')
        if aspects_data is None:
            return jsonify({'error': 'No aspects found'}), 404
            
        # Find and update the aspect
        aspect_found = False
        for index, row in aspects_data.iterrows():
            if int(row['id']) == aspect_id:
                aspects_data.at[index, 'nama'] = data['nama']
                aspects_data.at[index, 'updated_at'] = datetime.now().isoformat()
                aspect_found = True
                break
                
        if not aspect_found:
            return jsonify({'error': 'Aspect not found'}), 404
            
        # Save to storage (CSV for easier reading)
        success = storage_service.write_csv(aspects_data, 'config/aspects.csv')
        
        if success:
            return jsonify({'success': True}), 200
        else:
            return jsonify({'error': 'Failed to update aspect'}), 500
            
    except Exception as e:
        safe_print(f"Error updating aspect: {e}")
        return jsonify({'error': f'Failed to update aspect: {str(e)}'}), 500

@app.route('/api/config/aspects/<int:aspect_id>', methods=['DELETE'])
def delete_aspect(aspect_id):
    """Delete an aspect"""
    try:
        # Read existing aspects
        aspects_data = storage_service.read_csv('config/aspects.csv')
        if aspects_data is None:
            return jsonify({'error': 'No aspects found'}), 404
            
        # Filter out the aspect to delete
        aspects_data = aspects_data[aspects_data['id'] != aspect_id]
        
        # Save to storage (CSV for easier reading)
        success = storage_service.write_csv(aspects_data, 'config/aspects.csv')
        
        if success:
            return jsonify({'success': True}), 200
        else:
            return jsonify({'error': 'Failed to delete aspect'}), 500
            
    except Exception as e:
        safe_print(f"Error deleting aspect: {e}")
        return jsonify({'error': f'Failed to delete aspect: {str(e)}'}), 500

# CHECKLIST ENDPOINTS
@app.route('/api/config/checklist', methods=['GET'])
def get_checklist():
    """Get all checklist items with PIC assignments, optionally filtered by year"""
    from database import get_db_connection
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()

            # Filter by year if provided
            year = request.args.get('year')
            if year:
                try:
                    year_int = int(year)
                    cursor.execute("""
                        SELECT
                            c.id,
                            c.aspek,
                            c.deskripsi,
                            c.tahun,
                            c.created_at,
                            c.is_active,
                            a.subdirektorat
                        FROM checklist_gcg c
                        LEFT JOIN checklist_assignments a ON c.id = a.checklist_id AND c.tahun = a.tahun
                        WHERE c.tahun = ? AND c.is_active = 1
                        ORDER BY c.id
                    """, (year_int,))
                    safe_print(f"DEBUG: Fetching checklist for year {year_int}")
                except ValueError:
                    safe_print(f"WARNING: Invalid year parameter: {year}")
                    return jsonify({'checklist': []}), 200
            else:
                # Return all checklist items if no year specified
                cursor.execute("""
                    SELECT
                        c.id,
                        c.aspek,
                        c.deskripsi,
                        c.tahun,
                        c.created_at,
                        c.is_active,
                        a.subdirektorat
                    FROM checklist_gcg c
                    LEFT JOIN checklist_assignments a ON c.id = a.checklist_id AND c.tahun = a.tahun
                    WHERE c.is_active = 1
                    ORDER BY c.tahun DESC, c.id
                """)

            rows = cursor.fetchall()
            checklist_items = []
            for idx, row in enumerate(rows, 1):
                checklist_items.append({
                    'id': row[0],
                    'aspek': row[1] or '',
                    'deskripsi': row[2] or '',
                    'tahun': row[3],
                    'created_at': row[4],
                    'is_active': row[5],
                    'rowNumber': idx,
                    'pic': row[6] or ''  # PIC from checklist_assignments table
                })

            safe_print(f"DEBUG: Returning {len(checklist_items)} checklist items with PIC assignments")
            return jsonify({'checklist': checklist_items}), 200

    except Exception as e:
        safe_print(f"Error getting checklist: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'checklist': []}), 200

@app.route('/api/config/checklist', methods=['POST'])
def add_checklist():
    """Add a new checklist item to SQLite database"""
    from database import get_db_connection
    try:
        data = request.get_json()
        year = data.get('tahun')
        aspek = data.get('aspek', '')
        deskripsi = data.get('deskripsi', '')
        pic = data.get('pic', '')

        if not year or not deskripsi:
            return jsonify({'error': 'Year and description are required'}), 400

        with get_db_connection() as conn:
            cursor = conn.cursor()

            # Insert new checklist item
            cursor.execute("""
                INSERT INTO checklist_gcg (aspek, deskripsi, tahun, is_active)
                VALUES (?, ?, ?, ?)
            """, (aspek, deskripsi, year, True))

            checklist_id = cursor.lastrowid

            # If PIC is provided, create assignment
            if pic:
                cursor.execute("""
                    INSERT INTO checklist_assignments (checklist_id, subdirektorat, aspek, tahun)
                    VALUES (?, ?, ?, ?)
                """, (checklist_id, pic, aspek, year))

            # Prepare response
            checklist_data = {
                'id': checklist_id,
                'aspek': aspek,
                'deskripsi': deskripsi,
                'pic': pic,
                'tahun': year,
                'created_at': datetime.now().isoformat(),
                'is_active': True
            }

            safe_print(f"✅ Created checklist in SQLite: ID {checklist_id}, Year {year}")
            return jsonify(checklist_data), 201

    except Exception as e:
        safe_print(f"Error adding checklist: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'Failed to add checklist: {str(e)}'}), 500

@app.route('/api/config/checklist/<int:checklist_id>', methods=['PUT'])
def update_checklist(checklist_id):
    """Update an existing checklist item in SQLite and transfer files if PIC changes"""
    from database import get_db_connection
    try:
        data = request.get_json()

        # First, get existing data from database
        with get_db_connection() as conn:
            cursor = conn.cursor()

            # Get existing checklist item
            cursor.execute("""
                SELECT id, aspek, deskripsi, tahun FROM checklist_gcg WHERE id = ?
            """, (checklist_id,))
            existing = cursor.fetchone()

            if not existing:
                return jsonify({'error': 'Checklist item not found'}), 404

            # Get old PIC from assignments table
            cursor.execute("""
                SELECT subdirektorat FROM checklist_assignments
                WHERE checklist_id = ? AND tahun = ?
            """, (checklist_id, existing[3]))
            old_pic_row = cursor.fetchone()
            old_pic = old_pic_row[0] if old_pic_row else ''

        # Get new values from request (outside DB context to avoid closing issues)
        new_pic = data.get('pic', '')
        new_aspek = data.get('aspek', existing[1])
        new_deskripsi = data.get('deskripsi', existing[2])
        old_tahun = existing[3]
        new_tahun = data.get('tahun', old_tahun)

        # Debug logging
        safe_print(f"DEBUG: Checklist {checklist_id} PIC change check:")
        safe_print(f"  Old PIC: '{old_pic}'")
        safe_print(f"  New PIC: '{new_pic}'")
        safe_print(f"  Old Year: '{old_tahun}'")
        safe_print(f"  New Year: '{new_tahun}'")
        
        # Check if PIC or year is changing - this requires file transfer
        pic_changed = old_pic != new_pic
        year_changed = str(old_tahun) != str(new_tahun)
        files_transferred = False
        transfer_errors = []
        
        safe_print(f"  PIC changed: {pic_changed}")
        safe_print(f"  Year changed: {year_changed}")
        safe_print(f"  Both old and new PIC exist: {bool(old_pic and new_pic)}")
        
        # If PIC changes, transfer existing files from old location to new location
        if pic_changed and old_pic and new_pic:
            try:
                # Define old and new directory paths
                from werkzeug.utils import secure_filename
                old_pic_clean = secure_filename(old_pic.replace(' ', '_'))
                new_pic_clean = secure_filename(new_pic.replace(' ', '_'))
                old_dir = f"gcg-documents/{old_tahun}/{old_pic_clean}/{checklist_id}"
                new_dir = f"gcg-documents/{new_tahun}/{new_pic_clean}/{checklist_id}"

                safe_print(f"PIC change detected for checklist {checklist_id}")
                safe_print(f"Transferring files from: {old_dir}")
                safe_print(f"                    to: {new_dir}")

                # Transfer files between directories (local storage)
                import os
                import shutil

                # Use Path to get correct absolute path (backend is working directory)
                base_dir = Path(__file__).parent.parent / 'data'
                old_local_dir = base_dir / old_dir.replace('/', os.sep)
                new_local_dir = base_dir / new_dir.replace('/', os.sep)

                if old_local_dir.exists():
                    try:
                        # Create new directory structure
                        new_local_dir.parent.mkdir(parents=True, exist_ok=True)
                        # Move directory
                        shutil.move(str(old_local_dir), str(new_local_dir))
                        safe_print(f"✅ Successfully moved local directory from {old_local_dir} to {new_local_dir}")
                        files_transferred = True

                        # Update uploaded-files.xlsx tracking file with new paths
                        try:
                            files_data = storage_service.read_excel('uploaded-files.xlsx')
                            if files_data is not None and not files_data.empty:
                                # Find records for this checklist ID and year
                                mask = (files_data['checklistId'] == checklist_id) & (files_data['year'] == old_tahun)
                                if mask.any():
                                    # Update the PIC name and file paths for matching records
                                    for idx in files_data[mask].index:
                                        old_path = files_data.loc[idx, 'localFilePath']
                                        # Replace old PIC with new PIC in the path
                                        new_path = old_path.replace(f"/{old_pic_clean}/", f"/{new_pic_clean}/")
                                        # Also update year if it changed
                                        if year_changed:
                                            new_path = new_path.replace(f"gcg-documents/{old_tahun}/", f"gcg-documents/{new_tahun}/")

                                        files_data.loc[idx, 'localFilePath'] = new_path
                                        files_data.loc[idx, 'subdirektorat'] = new_pic
                                        if year_changed:
                                            files_data.loc[idx, 'year'] = new_tahun

                                        safe_print(f"📝 Updated file path: {old_path} → {new_path}")

                                    # Save updated tracking file
                                    storage_service.write_excel(files_data, 'uploaded-files.xlsx')
                                    safe_print(f"✅ Updated uploaded-files.xlsx with new paths")
                        except Exception as tracking_error:
                            safe_print(f"⚠️ Warning: Failed to update uploaded-files.xlsx: {tracking_error}")
                            # Don't fail the whole operation if tracking update fails

                        # CRITICAL: Also update file_path in SQLite database
                        try:
                            from database import get_db_connection
                            with get_db_connection() as conn:
                                cursor = conn.cursor()
                                # Get the new file path (same pattern as constructed path)
                                new_file_path = f"gcg-documents/{new_tahun}/{new_pic_clean}/{checklist_id}"
                                # Update the file_path for this checklist_id
                                cursor.execute("""
                                    UPDATE uploaded_files
                                    SET file_path = file_path || ''
                                    WHERE file_path LIKE ?
                                """, (f"%/{checklist_id}/%",))
                                # More precise update: replace old directory with new in file_path
                                cursor.execute("""
                                    UPDATE uploaded_files
                                    SET file_path = REPLACE(file_path, ?, ?)
                                    WHERE checklist_id = ? AND year = ?
                                """, (old_dir, new_dir, checklist_id, old_tahun if not year_changed else new_tahun))
                                conn.commit()
                                safe_print(f"✅ Updated database file_path for checklist_id {checklist_id}")
                        except Exception as db_error:
                            safe_print(f"⚠️ Warning: Failed to update database file_path: {db_error}")
                            # Don't fail the whole operation if database update fails

                    except Exception as move_error:
                        error_msg = f"Failed to move local directory: {str(move_error)}"
                        safe_print(f"❌ {error_msg}")
                        transfer_errors.append(error_msg)
                else:
                    safe_print(f"ℹ️ No local directory to transfer: {old_local_dir}")
                    files_transferred = True  # No directory to transfer is considered success

            except Exception as transfer_error:
                error_msg = f"General error during file transfer: {str(transfer_error)}"
                safe_print(f"❌ {error_msg}")
                transfer_errors.append(error_msg)

        # Now update the database in a new context
        with get_db_connection() as conn:
            cursor = conn.cursor()

            # Update the checklist item in database
            cursor.execute("""
                UPDATE checklist_gcg
                SET aspek = ?, deskripsi = ?, tahun = ?
                WHERE id = ?
            """, (new_aspek, new_deskripsi, new_tahun, checklist_id))

            # Update or create PIC assignment
            if new_pic:
                # Check if assignment exists
                cursor.execute("""
                    SELECT id FROM checklist_assignments
                    WHERE checklist_id = ? AND tahun = ?
                """, (checklist_id, new_tahun))
                assignment_exists = cursor.fetchone()

                if assignment_exists:
                    # Update existing assignment
                    cursor.execute("""
                        UPDATE checklist_assignments
                        SET subdirektorat = ?, aspek = ?
                        WHERE checklist_id = ? AND tahun = ?
                    """, (new_pic, new_aspek, checklist_id, new_tahun))
                else:
                    # Create new assignment
                    cursor.execute("""
                        INSERT INTO checklist_assignments (checklist_id, subdirektorat, aspek, tahun)
                        VALUES (?, ?, ?, ?)
                    """, (checklist_id, new_pic, new_aspek, new_tahun))
            elif old_pic:
                # PIC removed - delete assignment
                cursor.execute("""
                    DELETE FROM checklist_assignments
                    WHERE checklist_id = ? AND tahun = ?
                """, (checklist_id, new_tahun))

            # Prepare response
            response_data = {'success': True}
            if files_transferred:
                response_data['files_transferred'] = True
                response_data['message'] = f"Checklist updated and files transferred to new PIC directory"
            if transfer_errors:
                response_data['transfer_errors'] = transfer_errors
                response_data['warning'] = f"Checklist updated but some files failed to transfer: {len(transfer_errors)} errors"

            safe_print(f"✅ Updated checklist {checklist_id} in SQLite (PIC: {old_pic} → {new_pic})")
            return jsonify(response_data), 200

    except Exception as e:
        safe_print(f"Error updating checklist: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'Failed to update checklist: {str(e)}'}), 500

@app.route('/api/config/checklist/<int:checklist_id>', methods=['DELETE'])
def delete_checklist(checklist_id):
    """Delete a checklist item"""
    try:
        # Read existing checklist
        checklist_data = storage_service.read_csv('config/checklist.csv')
        if checklist_data is None:
            return jsonify({'error': 'No checklist found'}), 404
            
        # Filter out the checklist item to delete
        checklist_data = checklist_data[checklist_data['id'] != checklist_id]
        
        # Save to storage
        success = storage_service.write_csv(checklist_data, 'config/checklist.csv')
        
        if success:
            return jsonify({'success': True}), 200
        else:
            return jsonify({'error': 'Failed to delete checklist'}), 500
            
    except Exception as e:
        safe_print(f"Error deleting checklist: {e}")
        return jsonify({'error': f'Failed to delete checklist: {str(e)}'}), 500

@app.route('/api/check-files-exist/<int:checklist_id>', methods=['GET'])
def check_files_exist(checklist_id):
    """Check if files exist for a checklist item"""
    try:
        year = request.args.get('year', str(datetime.now().year))
        
        # Get checklist item to find current PIC
        checklist_data = storage_service.read_csv('config/checklist.csv')
        if checklist_data is None or checklist_data.empty:
            return jsonify({'hasFiles': False}), 200
        
        # Find the checklist item
        existing_item = checklist_data[checklist_data['id'] == checklist_id]
        if existing_item.empty:
            return jsonify({'hasFiles': False}), 200
        
        current_pic = existing_item.iloc[0].get('pic', '')
        if not current_pic:
            return jsonify({'hasFiles': False}), 200
        
        # Check if files exist in local storage
        import os
        from pathlib import Path
        from werkzeug.utils import secure_filename
        pic_clean = secure_filename(current_pic.replace(' ', '_'))
        directory_path = Path(__file__).parent.parent / 'data' / f"gcg-documents/{year}/{pic_clean}/{checklist_id}/"

        if directory_path.exists():
            # Filter out placeholder files and hidden files
            real_files = [
                f for f in directory_path.iterdir()
                if f.is_file() and
                not f.name.startswith('.') and
                not f.name.lower().startswith('placeholder') and
                f.stat().st_size > 0
            ]
            has_files = len(real_files) > 0
            safe_print(f"🔍 DEBUG: Checking local files for {directory_path}: found {len(list(directory_path.iterdir()))} items, {len(real_files)} real files")
            return jsonify({'hasFiles': has_files}), 200
        else:
            return jsonify({'hasFiles': False}), 200
    
    except Exception as e:
        safe_print(f"Error checking files existence: {e}")
        return jsonify({'hasFiles': False}), 200

@app.route('/api/config/checklist/clear', methods=['DELETE'])
def clear_checklist():
    """Clear all checklist data"""
    try:
        # Create empty DataFrame and save to clear the file
        empty_df = pd.DataFrame()
        success = storage_service.write_csv(empty_df, 'config/checklist.csv')
        
        if success:
            return jsonify({'success': True, 'message': 'Checklist data cleared'}), 200
        else:
            return jsonify({'error': 'Failed to clear checklist data'}), 500
            
    except Exception as e:
        safe_print(f"Error clearing checklist: {e}")
        return jsonify({'error': f'Failed to clear checklist: {str(e)}'}), 500

@app.route('/api/config/checklist/fix-ids', methods=['POST'])
def fix_checklist_ids():
    """Temporary endpoint to fix checklist IDs to proper year+row format"""
    try:
        # Read existing checklist
        existing_data = storage_service.read_csv('config/checklist.csv')
        if existing_data is None or existing_data.empty:
            return jsonify({'error': 'No checklist data found'}), 404
        
        safe_print(f"🔧 DEBUG: Fixing IDs for {len(existing_data)} checklist items")
        
        # Update each item with correct ID
        for index, row in existing_data.iterrows():
            year = int(row['tahun'])
            row_number = int(row['rowNumber'])
            correct_id = generate_checklist_id(year, row_number)
            existing_data.loc[index, 'id'] = correct_id
            safe_print(f"🔧 DEBUG: Row {row_number}: {row['id']} -> {correct_id}")
        
        # Save updated data
        success = storage_service.write_csv(existing_data, 'config/checklist.csv')
        
        if success:
            return jsonify({
                'success': True,
                'message': f'Successfully fixed {len(existing_data)} checklist IDs',
                'count': len(existing_data)
            }), 200
        else:
            return jsonify({'error': 'Failed to save fixed checklist to storage'}), 500
            
    except Exception as e:
        safe_print(f"Error fixing checklist IDs: {e}")
        return jsonify({'error': f'Failed to fix checklist IDs: {str(e)}'}), 500

@app.route('/api/config/checklist/batch', methods=['POST'])
def add_checklist_batch():
    """Add multiple checklist items in batch to both CSV and SQLite database"""
    from database import get_db_connection
    try:
        data = request.get_json()
        items = data.get('items', [])

        if not items:
            return jsonify({'error': 'No items provided'}), 400

        safe_print(f"📦 Batch adding {len(items)} checklist items")

        # Read existing checklist from CSV
        existing_data = storage_service.read_csv('config/checklist.csv')
        if existing_data is not None:
            checklist_df = existing_data
        else:
            checklist_df = pd.DataFrame()

        # Process each item in the batch
        batch_data = []
        for item in items:
            checklist_data = {
                'id': generate_checklist_id(item.get('tahun'), item.get('rowNumber')),
                'aspek': str(item.get('aspek', '')),
                'deskripsi': str(item.get('deskripsi', '')),
                'pic': str(item.get('pic', '')),
                'tahun': item.get('tahun'),
                'rowNumber': item.get('rowNumber'),
                'created_at': datetime.now().isoformat()
            }
            batch_data.append(checklist_data)

        # 1. Save to CSV file
        new_batch_df = pd.DataFrame(batch_data)
        updated_df = pd.concat([checklist_df, new_batch_df], ignore_index=True)
        csv_success = storage_service.write_csv(updated_df, 'config/checklist.csv')

        if not csv_success:
            return jsonify({'error': 'Failed to save checklist batch to CSV'}), 500

        # 2. Also save to SQLite database
        try:
            with get_db_connection() as conn:
                cursor = conn.cursor()
                db_saved_count = 0

                for item_data in batch_data:
                    try:
                        # Insert into checklist_gcg table
                        cursor.execute("""
                            INSERT INTO checklist_gcg (aspek, deskripsi, tahun, is_active, created_at)
                            VALUES (?, ?, ?, ?, ?)
                        """, (
                            item_data['aspek'],
                            item_data['deskripsi'],
                            item_data['tahun'],
                            1,  # is_active = True
                            item_data['created_at']
                        ))
                        db_saved_count += 1
                    except Exception as item_error:
                        safe_print(f"⚠️ Error saving item to SQLite: {item_error}")

                safe_print(f"✅ Saved {db_saved_count}/{len(batch_data)} items to SQLite database")
        except Exception as db_error:
            safe_print(f"⚠️ SQLite batch save error (CSV saved successfully): {db_error}")

        return jsonify({
            'success': True,
            'message': f'Successfully added {len(batch_data)} checklist items',
            'items': batch_data
        }), 201

    except Exception as e:
        safe_print(f"❌ Error adding checklist batch: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'Failed to add checklist batch: {str(e)}'}), 500

@app.route('/api/config/checklist/migrate-year', methods=['POST'])
def migrate_checklist_year():
    """Emergency endpoint to migrate checklist data from one year to another"""
    try:
        data = request.get_json()
        from_year = data.get('from_year')
        to_year = data.get('to_year')
        
        if not from_year or not to_year:
            return jsonify({'error': 'Both from_year and to_year are required'}), 400
        
        # Read existing checklist
        csv_data = storage_service.read_csv('config/checklist.csv')
        if csv_data is None:
            return jsonify({'error': 'No checklist data found'}), 404
        
        # Find items from source year
        source_items = csv_data[csv_data['tahun'] == from_year]
        if len(source_items) == 0:
            return jsonify({'error': f'No items found for year {from_year}'}), 404
        
        # Check if target year already has data
        target_items = csv_data[csv_data['tahun'] == to_year]
        if len(target_items) > 0:
            return jsonify({
                'error': f'Year {to_year} already has {len(target_items)} items. Migration aborted to prevent conflicts.',
                'existing_items': len(target_items)
            }), 409
        
        # Create migrated data with new year and IDs
        migrated_items = []
        for index, item in source_items.iterrows():
            # Generate new ID for target year
            new_id = generate_checklist_id(to_year, item.get('rowNumber'))
            
            migrated_item = {
                'id': new_id,
                'aspek': item.get('aspek', ''),
                'deskripsi': item.get('deskripsi', ''),
                'tahun': to_year,
                'rowNumber': item.get('rowNumber'),
                'pic': item.get('pic', ''),
                'created_at': datetime.now().isoformat()
            }
            migrated_items.append(migrated_item)
        
        # Add migrated items to existing data
        migrated_df = pd.DataFrame(migrated_items)
        updated_df = pd.concat([csv_data, migrated_df], ignore_index=True)
        
        # Save to storage
        success = storage_service.write_csv(updated_df, 'config/checklist.csv')
        
        if success:
            return jsonify({
                'success': True,
                'message': f'Successfully migrated {len(migrated_items)} items from year {from_year} to {to_year}',
                'migrated_count': len(migrated_items),
                'from_year': from_year,
                'to_year': to_year
            }), 200
        else:
            return jsonify({'error': 'Failed to save migrated data to storage'}), 500
            
    except Exception as e:
        safe_print(f"Error migrating checklist year: {e}")
        return jsonify({'error': f'Failed to migrate checklist year: {str(e)}'}), 500


# Register API route blueprints for SQLite backend
try:
    from api_routes import api_bp
    from api_config_routes import config_bp
    app.register_blueprint(api_bp)
    app.register_blueprint(config_bp)
    safe_print("✅ Registered SQLite API routes and config endpoints")
except ImportError as e:
    safe_print(f"⚠️ SQLite API routes not available: {e}")


# EXCEL EXPORT ENDPOINTS
from excel_exporter import ExcelExporter
exporter = ExcelExporter()

@app.route('/api/export/users', methods=['GET'])
def export_users_route():
    """Export users to Excel"""
    try:
        filepath = exporter.export_users()
        return send_file(filepath, as_attachment=True, download_name=os.path.basename(filepath))
    except Exception as e:
        safe_print(f"Error exporting users: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/export/checklist', methods=['GET'])
def export_checklist_route():
    """Export GCG checklist to Excel"""
    try:
        year = request.args.get('year', type=int)
        filepath = exporter.export_checklist_gcg(year=year)
        return send_file(filepath, as_attachment=True, download_name=os.path.basename(filepath))
    except Exception as e:
        safe_print(f"Error exporting checklist: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/export/documents', methods=['GET'])
def export_documents_route():
    """Export documents to Excel"""
    try:
        year = request.args.get('year', type=int)
        filepath = exporter.export_documents(year=year)
        return send_file(filepath, as_attachment=True, download_name=os.path.basename(filepath))
    except Exception as e:
        safe_print(f"Error exporting documents: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/export/org-structure', methods=['GET'])
def export_org_structure_route():
    """Export organizational structure to Excel"""
    try:
        year = request.args.get('year', type=int)
        filepath = exporter.export_organizational_structure(year=year)
        return send_file(filepath, as_attachment=True, download_name=os.path.basename(filepath))
    except ValueError as e:
        # Handle data not found errors with 404
        safe_print(f"Data not found: {e}")
        return jsonify({'error': str(e)}), 404
    except Exception as e:
        safe_print(f"Error exporting org structure: {e}")
        return jsonify({'error': f'Failed to export: {str(e)}'}), 500

@app.route('/api/export/gcg-assessment', methods=['GET'])
def export_gcg_assessment_route():
    """Export GCG assessment to Excel"""
    try:
        year = request.args.get('year', type=int)
        filepath = exporter.export_gcg_assessment(year=year)
        return send_file(filepath, as_attachment=True, download_name=os.path.basename(filepath))
    except Exception as e:
        safe_print(f"Error exporting GCG assessment: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/export/all', methods=['GET'])
def export_all_route():
    """Export all data to Excel"""
    try:
        year = request.args.get('year', type=int)
        filepath = exporter.export_all_data(year=year)
        return send_file(filepath, as_attachment=True, download_name=os.path.basename(filepath))
    except Exception as e:
        safe_print(f"Error exporting all data: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/export/history', methods=['GET'])
def export_history_route():
    """Get export history"""
    from database import get_db_connection
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT e.id, e.export_type, e.file_name, e.year,
                       u.name as exported_by_name, e.export_date,
                       e.row_count, e.file_size
                FROM excel_exports e
                LEFT JOIN users u ON e.exported_by = u.id
                ORDER BY e.export_date DESC
                LIMIT 50
            """)
            rows = cursor.fetchall()
            history = []
            for row in rows:
                history.append({
                    'id': row[0],
                    'export_type': row[1],
                    'file_name': row[2],
                    'year': row[3],
                    'exported_by_name': row[4],
                    'export_date': row[5],
                    'row_count': row[6],
                    'file_size': row[7]
                })
            return jsonify({'success': True, 'data': history}), 200
    except Exception as e:
        safe_print(f"Error getting export history: {e}")
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    safe_print(">> Starting POS Data Cleaner 2 Web API")
    safe_print(f"   Upload folder: {UPLOAD_FOLDER}")
    safe_print(f"   Output folder: {OUTPUT_FOLDER}")
    safe_print("   CORS enabled for React frontend")
    safe_print("   Production system integrated")
    safe_print("   Server starting on http://localhost:5000")
    
# =============================================================================
# PENGATURAN BARU CONFIGURATION ENDPOINTS
# =============================================================================

# DISABLED: Using blueprint route from api_config_routes.py instead
# @app.route('/api/config/tahun-buku', methods=['GET'])
# def get_tahun_buku():
#     """Get all tahun buku data"""
#     try:
#         # Read tahun buku from storage
#         tahun_data = storage_service.read_csv('config/tahun-buku.csv')
#
#         if tahun_data is None:
#             return jsonify({'tahun_buku': []}), 200
#
#         # Convert DataFrame to list of dictionaries
#         tahun_list = tahun_data.to_dict('records')
#
#         return jsonify({'tahun_buku': tahun_list}), 200
#
#     except Exception as e:
#         safe_print(f"❌ Error getting tahun buku: {e}")
#         return jsonify({'error': str(e)}), 500

# DISABLED: This route is now handled by api_config_routes.py blueprint
# The blueprint has proper cleanup logic for year reactivation

@app.route('/api/config/tahun-buku/<int:tahun_id>', methods=['DELETE'])
def delete_tahun_buku(tahun_id):
    """Delete a tahun buku by ID"""
    try:
        safe_print(f"🗑️ Deleting tahun buku with ID: {tahun_id}")
        
        # Read existing tahun buku data
        tahun_data = storage_service.read_csv('config/tahun-buku.csv')
        
        if tahun_data is None or tahun_data.empty:
            return jsonify({'error': 'No tahun buku data found'}), 404
            
        # Check if tahun exists
        if tahun_id not in tahun_data['id'].values:
            return jsonify({'error': 'Tahun buku not found'}), 404
        
        # Get the year value before deletion for cleanup
        year_to_delete = tahun_data[tahun_data['id'] == tahun_id]['tahun'].iloc[0]
        safe_print(f"🗑️ Deleting year: {year_to_delete}")
        
        # Remove the tahun buku entry
        filtered_data = tahun_data[tahun_data['id'] != tahun_id]
        
        # Save updated data to CSV
        success = storage_service.write_csv(filtered_data, 'config/tahun-buku.csv')

        if success:
            # IMPORTANT: Also soft-delete from database years table
            try:
                from database import get_db_connection
                with get_db_connection() as conn:
                    cursor = conn.cursor()
                    cursor.execute("UPDATE years SET is_active = 0 WHERE year = ?", (year_to_delete,))
                    print(f"[OK] Soft-deleted year {year_to_delete} from database years table")
            except Exception as db_error:
                print(f"[WARNING] Could not soft-delete from database: {db_error}")

            safe_print(f"✅ Successfully deleted tahun buku {tahun_id} (year {year_to_delete})")

            # Clean up all related data for this year
            safe_print(f"🧹 Cleaning up all related data for year {year_to_delete}")
            cleanup_stats = {}

            try:
                # 1. Clean up checklist data
                checklist_data = storage_service.read_csv('config/checklist.csv')
                if checklist_data is not None and not checklist_data.empty:
                    original_count = len(checklist_data)
                    checklist_data = checklist_data[checklist_data['tahun'] != year_to_delete]
                    storage_service.write_csv(checklist_data, 'config/checklist.csv')
                    cleanup_stats['checklist'] = original_count - len(checklist_data)
                    safe_print(f"  ✅ Cleaned {cleanup_stats['checklist']} checklist items")

                # 2. Clean up aspects data
                aspects_data = storage_service.read_csv('config/aspects.csv')
                if aspects_data is not None and not aspects_data.empty:
                    original_count = len(aspects_data)
                    aspects_data = aspects_data[aspects_data['tahun'] != year_to_delete]
                    storage_service.write_csv(aspects_data, 'config/aspects.csv')
                    cleanup_stats['aspects'] = original_count - len(aspects_data)
                    safe_print(f"  ✅ Cleaned {cleanup_stats['aspects']} aspects")

                # 3. Clean up struktur organisasi data
                struktur_data = storage_service.read_csv('config/struktur-organisasi.csv')
                if struktur_data is not None and not struktur_data.empty:
                    original_count = len(struktur_data)
                    struktur_data = struktur_data[struktur_data['tahun'] != year_to_delete]
                    storage_service.write_csv(struktur_data, 'config/struktur-organisasi.csv')
                    cleanup_stats['struktur'] = original_count - len(struktur_data)
                    safe_print(f"  ✅ Cleaned {cleanup_stats['struktur']} struktur organisasi items")

                # 4. Clean up AOI tables
                aoi_tables_data = storage_service.read_csv('config/aoi-tables.csv')
                if aoi_tables_data is not None and not aoi_tables_data.empty:
                    original_count = len(aoi_tables_data)
                    aoi_tables_data = aoi_tables_data[aoi_tables_data['tahun'] != year_to_delete]
                    storage_service.write_csv(aoi_tables_data, 'config/aoi-tables.csv')
                    cleanup_stats['aoi_tables'] = original_count - len(aoi_tables_data)
                    safe_print(f"  ✅ Cleaned {cleanup_stats['aoi_tables']} AOI tables")

                # 5. Clean up AOI recommendations (if they have year field)
                try:
                    aoi_recs_data = storage_service.read_csv('config/aoi-recommendations.csv')
                    if aoi_recs_data is not None and not aoi_recs_data.empty and 'tahun' in aoi_recs_data.columns:
                        original_count = len(aoi_recs_data)
                        aoi_recs_data = aoi_recs_data[aoi_recs_data['tahun'] != year_to_delete]
                        storage_service.write_csv(aoi_recs_data, 'config/aoi-recommendations.csv')
                        cleanup_stats['aoi_recommendations'] = original_count - len(aoi_recs_data)
                        safe_print(f"  ✅ Cleaned {cleanup_stats['aoi_recommendations']} AOI recommendations")
                except Exception as e:
                    safe_print(f"  ⚠️ AOI recommendations cleanup skipped: {e}")

                # 6. Clean up uploaded files tracking
                uploaded_files_data = storage_service.read_excel('uploaded-files.xlsx')
                if uploaded_files_data is not None and not uploaded_files_data.empty:
                    original_count = len(uploaded_files_data)
                    uploaded_files_data = uploaded_files_data[uploaded_files_data['year'] != year_to_delete]
                    storage_service.write_excel(uploaded_files_data, 'uploaded-files.xlsx')
                    cleanup_stats['uploaded_files'] = original_count - len(uploaded_files_data)
                    safe_print(f"  ✅ Cleaned {cleanup_stats['uploaded_files']} uploaded file records")

                # 7. Clean up checklist assignments
                try:
                    assignments_data = storage_service.read_csv('config/checklist-assignments.csv')
                    if assignments_data is not None and not assignments_data.empty and 'year' in assignments_data.columns:
                        original_count = len(assignments_data)
                        assignments_data = assignments_data[assignments_data['year'] != year_to_delete]
                        storage_service.write_csv(assignments_data, 'config/checklist-assignments.csv')
                        cleanup_stats['assignments'] = original_count - len(assignments_data)
                        safe_print(f"  ✅ Cleaned {cleanup_stats['assignments']} checklist assignments")
                except Exception as e:
                    safe_print(f"  ⚠️ Assignments cleanup skipped: {e}")

                # 8. Clean up users with year-specific data
                users_data = storage_service.read_csv('config/users.csv')
                if users_data is not None and not users_data.empty and 'tahun' in users_data.columns:
                    original_count = len(users_data)
                    users_data = users_data[users_data['tahun'] != year_to_delete]
                    storage_service.write_csv(users_data, 'config/users.csv')
                    cleanup_stats['users'] = original_count - len(users_data)
                    safe_print(f"  ✅ Cleaned {cleanup_stats['users']} year-specific users")

                # 9. Clean up AOI documents tracking
                try:
                    aoi_docs_data = storage_service.read_csv('aoi-documents.csv')
                    if aoi_docs_data is not None and not aoi_docs_data.empty and 'year' in aoi_docs_data.columns:
                        original_count = len(aoi_docs_data)
                        aoi_docs_data = aoi_docs_data[aoi_docs_data['year'] != year_to_delete]
                        storage_service.write_csv(aoi_docs_data, 'aoi-documents.csv')
                        cleanup_stats['aoi_documents'] = original_count - len(aoi_docs_data)
                        safe_print(f"  ✅ Cleaned {cleanup_stats['aoi_documents']} AOI document records")
                except Exception as e:
                    safe_print(f"  ⚠️ AOI documents cleanup skipped: {e}")

                # 10. Clean up SQLite database tables
                safe_print(f"🗄️ Cleaning up SQLite database tables for year {year_to_delete}")
                try:
                    from database import get_db_connection
                    with get_db_connection() as conn:
                        cursor = conn.cursor()

                        # Clean struktur organisasi tables (order matters - delete children first)
                        cursor.execute("DELETE FROM divisi WHERE tahun = ?", (year_to_delete,))
                        divisi_deleted = cursor.rowcount

                        cursor.execute("DELETE FROM subdirektorat WHERE tahun = ?", (year_to_delete,))
                        subdirektorat_deleted = cursor.rowcount

                        cursor.execute("DELETE FROM direktorat WHERE tahun = ?", (year_to_delete,))
                        direktorat_deleted = cursor.rowcount

                        cursor.execute("DELETE FROM anak_perusahaan WHERE tahun = ?", (year_to_delete,))
                        anak_perusahaan_deleted = cursor.rowcount

                        # Clean users table - NOTE: users table doesn't have tahun column
                        # Users are cleaned from CSV file only (see step 8 below)
                        # cursor.execute("DELETE FROM users WHERE tahun = ?", (year_to_delete,))
                        users_deleted = 0  # Users cleaned from CSV, not database

                        # Clean checklist_gcg table
                        cursor.execute("DELETE FROM checklist_gcg WHERE tahun = ?", (year_to_delete,))
                        checklist_gcg_deleted = cursor.rowcount

                        conn.commit()

                        cleanup_stats['db_divisi'] = divisi_deleted
                        cleanup_stats['db_subdirektorat'] = subdirektorat_deleted
                        cleanup_stats['db_direktorat'] = direktorat_deleted
                        cleanup_stats['db_anak_perusahaan'] = anak_perusahaan_deleted
                        cleanup_stats['db_users'] = users_deleted
                        cleanup_stats['db_checklist_gcg'] = checklist_gcg_deleted

                        safe_print(f"  ✅ Database cleanup completed:")
                        safe_print(f"     - {divisi_deleted} divisi records")
                        safe_print(f"     - {subdirektorat_deleted} subdirektorat records")
                        safe_print(f"     - {direktorat_deleted} direktorat records")
                        safe_print(f"     - {anak_perusahaan_deleted} anak_perusahaan records")
                        safe_print(f"     - {users_deleted} users records")
                        safe_print(f"     - {checklist_gcg_deleted} checklist_gcg records")

                except Exception as db_error:
                    safe_print(f"  ⚠️ Database cleanup error: {db_error}")
                    cleanup_stats['db_error'] = str(db_error)

                safe_print(f"✅ All related data cleaned up successfully")

            except Exception as cleanup_error:
                safe_print(f"⚠️ Error during cleanup: {cleanup_error}")
                cleanup_stats['error'] = str(cleanup_error)

            return jsonify({
                'success': True,
                'message': f'Tahun buku {year_to_delete} and all related data deleted successfully',
                'deleted_year': int(year_to_delete),
                'cleanup_stats': cleanup_stats
            }), 200
        else:
            return jsonify({'error': 'Failed to delete tahun buku'}), 500
            
    except Exception as e:
        safe_print(f"❌ Error deleting tahun buku: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/config/struktur-organisasi', methods=['GET'])
def get_struktur_organisasi():
    """Get all struktur organisasi data from SQLite, optionally filtered by year"""
    from database import get_db_connection
    try:
        year = request.args.get('year', type=int)
        safe_print(f"📋 GET struktur-organisasi - year filter: {year}")

        with get_db_connection() as conn:
            cursor = conn.cursor()

            # Get direktorat with optional year filter
            if year:
                cursor.execute("""
                    SELECT id, nama, deskripsi, tahun, created_at, is_active
                    FROM direktorat
                    WHERE is_active = 1 AND tahun = ?
                    ORDER BY nama
                """, (year,))
            else:
                cursor.execute("""
                    SELECT id, nama, deskripsi, tahun, created_at, is_active
                    FROM direktorat
                    WHERE is_active = 1
                    ORDER BY nama
                """)
            direktorat = []
            for row in cursor.fetchall():
                direktorat.append({
                    'id': row[0],
                    'nama': row[1],
                    'deskripsi': row[2] or '',
                    'tahun': row[3],
                    'createdAt': row[4],
                    'isActive': row[5],
                    'type': 'direktorat'
                })

            # Get subdirektorat with optional year filter
            if year:
                cursor.execute("""
                    SELECT id, nama, direktorat_id, deskripsi, tahun, created_at, is_active
                    FROM subdirektorat
                    WHERE is_active = 1 AND tahun = ?
                    ORDER BY nama
                """, (year,))
            else:
                cursor.execute("""
                    SELECT id, nama, direktorat_id, deskripsi, tahun, created_at, is_active
                    FROM subdirektorat
                    WHERE is_active = 1
                    ORDER BY nama
                """)
            subdirektorat = []
            for row in cursor.fetchall():
                subdirektorat.append({
                    'id': row[0],
                    'nama': row[1],
                    'direktoratId': row[2],
                    'parent_id': row[2],
                    'deskripsi': row[3] or '',
                    'tahun': row[4],
                    'createdAt': row[5],
                    'isActive': row[6],
                    'type': 'subdirektorat'
                })

            # Get divisi with optional year filter
            if year:
                cursor.execute("""
                    SELECT id, nama, subdirektorat_id, deskripsi, tahun, created_at, is_active
                    FROM divisi
                    WHERE is_active = 1 AND tahun = ?
                    ORDER BY nama
                """, (year,))
            else:
                cursor.execute("""
                    SELECT id, nama, subdirektorat_id, deskripsi, tahun, created_at, is_active
                    FROM divisi
                    WHERE is_active = 1
                    ORDER BY nama
                """)
            divisi = []
            for row in cursor.fetchall():
                divisi.append({
                    'id': row[0],
                    'nama': row[1],
                    'subdirektoratId': row[2],
                    'parent_id': row[2],
                    'deskripsi': row[3] or '',
                    'tahun': row[4],
                    'createdAt': row[5],
                    'isActive': row[6],
                    'type': 'divisi'
                })

            # Get anak perusahaan with optional year filter
            if year:
                cursor.execute("""
                    SELECT id, nama, kategori, deskripsi, tahun, created_at, is_active
                    FROM anak_perusahaan
                    WHERE is_active = 1 AND tahun = ?
                    ORDER BY nama
                """, (year,))
            else:
                cursor.execute("""
                    SELECT id, nama, kategori, deskripsi, tahun, created_at, is_active
                    FROM anak_perusahaan
                    WHERE is_active = 1
                    ORDER BY nama
                """)
            anak_perusahaan = []
            for row in cursor.fetchall():
                anak_perusahaan.append({
                    'id': row[0],
                    'nama': row[1],
                    'kategori': row[2],
                    'deskripsi': row[3] or '',
                    'tahun': row[4],
                    'createdAt': row[5],
                    'isActive': row[6],
                    'type': 'anak_perusahaan'
                })

            result = {
                'direktorat': direktorat,
                'subdirektorat': subdirektorat,
                'divisi': divisi,
                'anak_perusahaan': anak_perusahaan
            }

            safe_print(f"✓ Returning struktur organisasi: {len(direktorat)} dir, {len(subdirektorat)} sub, {len(divisi)} div, {len(anak_perusahaan)} anak")
            return jsonify(result), 200

    except Exception as e:
        safe_print(f"❌ Error getting struktur organisasi: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/config/struktur-organisasi', methods=['POST'])
def add_struktur_organisasi():
    """Add a new struktur organisasi item to SQLite database"""
    from database import get_db_connection
    try:
        data = request.get_json()

        required_fields = ['type', 'nama', 'tahun']
        for field in required_fields:
            if not data.get(field):
                return jsonify({'error': f'{field} is required'}), 400

        # Validate type
        valid_types = ['direktorat', 'subdirektorat', 'divisi', 'anak_perusahaan']
        if data['type'] not in valid_types:
            return jsonify({'error': f'type must be one of: {", ".join(valid_types)}'}), 400

        item_type = data['type']
        nama = data['nama']
        deskripsi = data.get('deskripsi', '')
        tahun = data['tahun']
        parent_id = data.get('parent_id')

        safe_print(f"➕ Creating {item_type}: {nama} (tahun: {tahun})")

        # Save to SQLite database
        with get_db_connection() as conn:
            cursor = conn.cursor()

            if item_type == 'direktorat':
                cursor.execute("""
                    INSERT INTO direktorat (nama, deskripsi, tahun, is_active, created_at)
                    VALUES (?, ?, ?, ?, ?)
                """, (nama, deskripsi, tahun, 1, datetime.now().isoformat()))
                new_id = cursor.lastrowid

            elif item_type == 'subdirektorat':
                cursor.execute("""
                    INSERT INTO subdirektorat (nama, direktorat_id, deskripsi, tahun, is_active, created_at)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (nama, parent_id, deskripsi, tahun, 1, datetime.now().isoformat()))
                new_id = cursor.lastrowid

            elif item_type == 'divisi':
                cursor.execute("""
                    INSERT INTO divisi (nama, subdirektorat_id, deskripsi, tahun, is_active, created_at)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (nama, parent_id, deskripsi, tahun, 1, datetime.now().isoformat()))
                new_id = cursor.lastrowid

            elif item_type == 'anak_perusahaan':
                # Default to 'Anak Perusahaan' if not provided
                kategori = data.get('kategori', 'Anak Perusahaan')
                if not kategori or kategori.strip() == '':
                    kategori = 'Anak Perusahaan'
                cursor.execute("""
                    INSERT INTO anak_perusahaan (nama, kategori, deskripsi, tahun, is_active, created_at)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (nama, kategori, deskripsi, tahun, 1, datetime.now().isoformat()))
                new_id = cursor.lastrowid

        new_struktur = {
            'id': new_id,
            'type': item_type,
            'nama': nama,
            'deskripsi': deskripsi,
            'tahun': tahun,
            'parent_id': parent_id,
            'created_at': datetime.now().isoformat()
        }

        safe_print(f"✅ Created {item_type} with ID: {new_id}")
        return jsonify({'success': True, 'struktur': new_struktur}), 201

    except Exception as e:
        safe_print(f"❌ Error adding struktur organisasi: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/config/struktur-organisasi/batch', methods=['POST'])
def add_struktur_organisasi_batch():
    """Add multiple struktur organisasi items in a single transaction with proper ID mapping"""
    from database import get_db_connection
    try:
        data = request.get_json()
        items = data.get('items', [])

        safe_print(f"🔥 BATCH API CALLED: Received {len(items)} items")
        if items:
            safe_print(f"🔥 Sample item: {items[0]}")

        if not items:
            safe_print("❌ No items provided in batch request")
            return jsonify({'error': 'No items provided'}), 400

        # Sort items by dependency order: direktorat first, then subdirektorat, then divisi
        type_order = {'direktorat': 0, 'subdirektorat': 1, 'divisi': 2, 'anak_perusahaan': 3}
        items.sort(key=lambda x: type_order.get(x.get('type', ''), 999))

        # Track original ID to new ID mappings for hierarchical relationships
        id_mappings = {}
        created_items = []  # For response

        with get_db_connection() as conn:
            cursor = conn.cursor()

            for item in items:
                item_type = item.get('type')
                nama = item.get('nama')
                deskripsi = item.get('deskripsi', '')
                tahun = item.get('tahun', datetime.now().year)

                # Map parent_id if this item has a parent
                mapped_parent_id = None
                if item.get('parent_original_id'):
                    mapped_parent_id = id_mappings.get(item['parent_original_id'])
                    if not mapped_parent_id:
                        safe_print(f"Warning: Could not find mapping for parent_original_id {item['parent_original_id']}")
                elif item.get('parent_id'):
                    # Direct parent_id (for backward compatibility)
                    mapped_parent_id = item['parent_id']

                # Insert into appropriate table
                if item_type == 'direktorat':
                    cursor.execute("""
                        INSERT INTO direktorat (nama, deskripsi, tahun, created_at, is_active)
                        VALUES (?, ?, ?, ?, 1)
                    """, (nama, deskripsi, tahun, datetime.now().isoformat()))
                    new_id = cursor.lastrowid

                elif item_type == 'subdirektorat':
                    cursor.execute("""
                        INSERT INTO subdirektorat (nama, direktorat_id, deskripsi, tahun, created_at, is_active)
                        VALUES (?, ?, ?, ?, ?, 1)
                    """, (nama, mapped_parent_id, deskripsi, tahun, datetime.now().isoformat()))
                    new_id = cursor.lastrowid

                elif item_type == 'divisi':
                    cursor.execute("""
                        INSERT INTO divisi (nama, subdirektorat_id, deskripsi, tahun, created_at, is_active)
                        VALUES (?, ?, ?, ?, ?, 1)
                    """, (nama, mapped_parent_id, deskripsi, tahun, datetime.now().isoformat()))
                    new_id = cursor.lastrowid

                elif item_type == 'anak_perusahaan':
                    # Default to 'Anak Perusahaan' if not provided
                    kategori = item.get('kategori', 'Anak Perusahaan')
                    if not kategori or kategori.strip() == '':
                        kategori = 'Anak Perusahaan'
                    cursor.execute("""
                        INSERT INTO anak_perusahaan (nama, kategori, deskripsi, tahun, created_at, is_active)
                        VALUES (?, ?, ?, ?, ?, 1)
                    """, (nama, kategori, deskripsi, tahun, datetime.now().isoformat()))
                    new_id = cursor.lastrowid
                else:
                    continue

                # Store mapping from original_id to new_id
                if item.get('original_id'):
                    id_mappings[item['original_id']] = new_id

                created_items.append({
                    'id': new_id,
                    'nama': nama,
                    'deskripsi': deskripsi,
                    'tahun': tahun,
                    'type': item_type,
                    'parent_id': mapped_parent_id,
                    'created_at': datetime.now().isoformat()
                })

                safe_print(f"✓ Created {item_type}: {nama} (ID: {new_id}, Parent: {mapped_parent_id}, Year: {tahun})")

            conn.commit()

        safe_print(f"✅ Batch saved {len(created_items)} struktur organisasi items to SQLite!")
        return jsonify({
            'success': True,
            'added_count': len(created_items),
            'created': created_items,
            'id_mappings': id_mappings
        }), 201

    except Exception as e:
        safe_print(f"❌ Error adding batch struktur organisasi: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

# DISABLED: This route conflicts with blueprint route and uses CSV instead of SQLite
# Blueprint route at api_config_routes.py:537 handles DELETE correctly with SQLite (soft delete)
# @app.route('/api/config/struktur-organisasi/<int:struktur_id>', methods=['DELETE'])
# def delete_struktur_organisasi(struktur_id):
#     """Delete a struktur organisasi item"""
#     ... (disabled to use SQLite version from blueprint)
# @app.route('/api/config/struktur-organisasi/<int:struktur_id>', methods=['PUT'])
# def update_struktur_organisasi(struktur_id):
#     """Update a struktur organisasi item"""
#     ... (disabled to prevent conflict)

# CHECKLIST ASSIGNMENTS ENDPOINTS
@app.route('/api/config/assignments', methods=['GET'])
def get_assignments():
    """Get all checklist assignments"""
    try:
        assignments_data = storage_service.read_csv('config/checklist-assignments.csv')
        if assignments_data is not None:
            assignments = assignments_data.to_dict(orient='records')
            return jsonify({'assignments': assignments}), 200
        else:
            return jsonify({'assignments': []}), 200
    except Exception as e:
        safe_print(f"❌ Error getting assignments: {e}")
        return jsonify({'assignments': []}), 200

@app.route('/api/config/assignments', methods=['POST'])
def add_assignment():
    """Add or update checklist assignment"""
    try:
        data = request.get_json()
        
        assignment_data = {
            'checklistId': data.get('checklistId'),
            'assignedTo': data.get('assignedTo'),
            'assignmentType': data.get('assignmentType'),  # 'divisi' or 'subdirektorat'
            'year': data.get('year'),
            'createdAt': datetime.now().isoformat()
        }
        
        # Read existing assignments
        existing_data = storage_service.read_csv('config/checklist-assignments.csv')
        if existing_data is not None and not existing_data.empty:
            assignments_df = existing_data
            # Remove existing assignment for this checklistId if exists
            assignments_df = assignments_df[assignments_df['checklistId'] != assignment_data['checklistId']]
        else:
            # Create empty DataFrame with proper columns
            assignments_df = pd.DataFrame()
        
        # Add new assignment
        new_assignment_df = pd.DataFrame([assignment_data])
        updated_df = pd.concat([assignments_df, new_assignment_df], ignore_index=True)
        
        # Save to storage
        success = storage_service.write_csv(updated_df, 'config/checklist-assignments.csv')
        
        if success:
            return jsonify(assignment_data), 201
        else:
            return jsonify({'error': 'Failed to save assignment'}), 500
            
    except Exception as e:
        safe_print(f"❌ Error adding assignment: {e}")
        return jsonify({'error': f'Failed to add assignment: {str(e)}'}), 500

@app.route('/api/config/assignments/<int:checklist_id>', methods=['DELETE'])
def delete_assignment(checklist_id):
    """Delete assignment for a checklist item"""
    try:
        assignments_data = storage_service.read_csv('config/checklist-assignments.csv')
        if assignments_data is None:
            return jsonify({'success': True}), 200
            
        # Filter out the assignment to delete
        assignments_data = assignments_data[assignments_data['checklistId'] != checklist_id]
        
        # Save to storage
        success = storage_service.write_csv(assignments_data, 'config/checklist-assignments.csv')
        
        if success:
            return jsonify({'success': True}), 200
        else:
            return jsonify({'error': 'Failed to delete assignment'}), 500
            
    except Exception as e:
        safe_print(f"❌ Error deleting assignment: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/check-row-files/<int:year>/<pic_name>/<int:row_number>', methods=['GET'])
def check_row_files(year, pic_name, row_number):
    """Check if files exist for a specific row"""
    try:
        path = f"gcg-documents/{year}/{pic_name}/{row_number}"
        
        # Check if directory exists
        files = storage_service.list_files(path)
        has_files = len(files) > 0
        
        return jsonify({
            'success': True,
            'hasFiles': has_files,
            'fileCount': len(files),
            'files': files
        }), 200
        
    except Exception as e:
        safe_print(f"❌ Error checking row files: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/delete-row-files/<int:year>/<pic_name>/<int:row_number>', methods=['DELETE'])
def delete_row_files(year, pic_name, row_number):
    """Delete all files for a specific row"""
    try:
        path = f"gcg-documents/{year}/{pic_name}/{row_number}"
        
        # List files first
        files = storage_service.list_files(path)
        
        if not files:
            return jsonify({
                'success': True,
                'message': 'No files to delete',
                'deletedCount': 0
            }), 200
        
        # Delete all files in the directory
        deleted_count = 0
        for file_info in files:
            file_path = f"{path}/{file_info['name']}"
            success = storage_service.delete_file(file_path)
            if success:
                deleted_count += 1
        
        return jsonify({
            'success': True,
            'message': f'Successfully deleted {deleted_count} files',
            'deletedCount': deleted_count,
            'totalFiles': len(files)
        }), 200
        
    except Exception as e:
        safe_print(f"❌ Error deleting row files: {e}")
        return jsonify({'error': str(e)}), 500

# Bulk delete endpoints for year data management
@app.route('/api/bulk-delete/<int:year>/preview', methods=['GET'])
def preview_bulk_delete(year):
    """Preview what data would be deleted for a specific year"""
    try:
        safe_print(f"📋 Previewing bulk delete for year {year}")
        
        # Initialize counters
        preview_data = {
            'year': year,
            'checklist_items': 0,
            'aspects': 0,
            'users': 0,
            'organizational_data': {
                'direktorat': 0,
                'subdirektorat': 0,
                'divisi': 0
            },
            'uploaded_files': 0,
            'total_items': 0
        }
        
        # Count checklist items for the year
        try:
            checklist_data = storage_service.read_csv('config/checklist.csv')
            if checklist_data is not None:
                year_checklist = checklist_data[checklist_data['tahun'] == year]
                preview_data['checklist_items'] = len(year_checklist)
        except Exception as e:
            safe_print(f"⚠️ Error counting checklist items: {e}")
        
        # Count aspects for the year
        try:
            aspects_data = storage_service.read_csv('config/aspects.csv')
            if aspects_data is not None:
                year_aspects = aspects_data[aspects_data['tahun'] == year]
                preview_data['aspects'] = len(year_aspects)
        except Exception as e:
            safe_print(f"⚠️ Error counting aspects: {e}")
        
        # Count users for the year
        try:
            users_data = storage_service.read_csv('config/users.csv')
            if users_data is not None:
                year_users = users_data[users_data['tahun'] == year]
                preview_data['users'] = len(year_users)
        except Exception as e:
            safe_print(f"⚠️ Error counting users: {e}")
        
        # Count organizational data for the year
        try:
            org_data = storage_service.read_csv('config/struktur-organisasi.csv')
            if org_data is not None:
                year_org = org_data[org_data['tahun'] == year]
                preview_data['organizational_data']['direktorat'] = len(year_org[year_org['jenis'] == 'direktorat'])
                preview_data['organizational_data']['subdirektorat'] = len(year_org[year_org['jenis'] == 'subdirektorat'])
                preview_data['organizational_data']['divisi'] = len(year_org[year_org['jenis'] == 'divisi'])
        except Exception as e:
            safe_print(f"⚠️ Error counting organizational data: {e}")
        
        # Count uploaded files for the year
        try:
            files = storage_service.list_files(f"gcg-documents/{year}")
            if files:
                # Count only actual files, not directories
                file_count = 0
                for file_info in files:
                    if not file_info.get('name', '').endswith('/'):
                        file_count += 1
                preview_data['uploaded_files'] = file_count
        except Exception as e:
            safe_print(f"⚠️ Error counting uploaded files: {e}")
        
        # Calculate total items
        preview_data['total_items'] = (
            preview_data['checklist_items'] +
            preview_data['aspects'] +
            preview_data['users'] +
            preview_data['organizational_data']['direktorat'] +
            preview_data['organizational_data']['subdirektorat'] +
            preview_data['organizational_data']['divisi'] +
            preview_data['uploaded_files']
        )
        
        safe_print(f"✅ Preview complete for year {year}: {preview_data['total_items']} total items")
        return jsonify(preview_data), 200
        
    except Exception as e:
        safe_print(f"❌ Error previewing bulk delete: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/bulk-delete/<int:year>', methods=['DELETE'])
def bulk_delete_year_data(year):
    """Delete all data for a specific year"""
    try:
        safe_print(f"🗑️ Starting bulk delete for year {year}")
        
        deleted_summary = {
            'year': year,
            'checklist_items': 0,
            'aspects': 0,
            'users': 0,
            'organizational_data': {
                'direktorat': 0,
                'subdirektorat': 0,
                'divisi': 0
            },
            'uploaded_files': 0,
            'assignments': 0
        }
        
        # 1. Delete checklist items for the year
        try:
            checklist_data = storage_service.read_csv('config/checklist.csv')
            if checklist_data is not None:
                original_count = len(checklist_data)
                year_checklist = checklist_data[checklist_data['tahun'] == year]
                deleted_summary['checklist_items'] = len(year_checklist)
                
                # Keep only items not from this year
                remaining_checklist = checklist_data[checklist_data['tahun'] != year]
                success = storage_service.write_csv(remaining_checklist, 'config/checklist.csv')
                if success:
                    safe_print(f"✅ Deleted {deleted_summary['checklist_items']} checklist items for year {year}")
                else:
                    safe_print(f"❌ Failed to delete checklist items for year {year}")
        except Exception as e:
            safe_print(f"⚠️ Error deleting checklist items: {e}")
        
        # 2. Delete aspects for the year
        try:
            aspects_data = storage_service.read_csv('config/aspects.csv')
            if aspects_data is not None:
                year_aspects = aspects_data[aspects_data['tahun'] == year]
                deleted_summary['aspects'] = len(year_aspects)
                
                # Keep only aspects not from this year
                remaining_aspects = aspects_data[aspects_data['tahun'] != year]
                success = storage_service.write_csv(remaining_aspects, 'config/aspects.csv')
                if success:
                    safe_print(f"✅ Deleted {deleted_summary['aspects']} aspects for year {year}")
        except Exception as e:
            safe_print(f"⚠️ Error deleting aspects: {e}")
        
        # 3. Delete users for the year
        try:
            users_data = storage_service.read_csv('config/users.csv')
            if users_data is not None:
                year_users = users_data[users_data['tahun'] == year]
                deleted_summary['users'] = len(year_users)
                
                # Keep only users not from this year
                remaining_users = users_data[users_data['tahun'] != year]
                success = storage_service.write_csv(remaining_users, 'config/users.csv')
                if success:
                    safe_print(f"✅ Deleted {deleted_summary['users']} users for year {year}")
        except Exception as e:
            safe_print(f"⚠️ Error deleting users: {e}")
        
        # 4. Delete organizational data for the year
        try:
            org_data = storage_service.read_csv('config/struktur-organisasi.csv')
            if org_data is not None:
                year_org = org_data[org_data['tahun'] == year]
                deleted_summary['organizational_data']['direktorat'] = len(year_org[year_org['jenis'] == 'direktorat'])
                deleted_summary['organizational_data']['subdirektorat'] = len(year_org[year_org['jenis'] == 'subdirektorat'])
                deleted_summary['organizational_data']['divisi'] = len(year_org[year_org['jenis'] == 'divisi'])
                
                # Keep only organizational data not from this year
                remaining_org = org_data[org_data['tahun'] != year]
                success = storage_service.write_csv(remaining_org, 'config/struktur-organisasi.csv')
                if success:
                    total_org_deleted = sum(deleted_summary['organizational_data'].values())
                    safe_print(f"✅ Deleted {total_org_deleted} organizational items for year {year}")
        except Exception as e:
            safe_print(f"⚠️ Error deleting organizational data: {e}")
        
        # 5. Delete assignments for the year
        try:
            assignments_data = storage_service.read_csv('config/checklist-assignments.csv')
            if assignments_data is not None:
                year_assignments = assignments_data[assignments_data['tahun'] == year]
                deleted_summary['assignments'] = len(year_assignments)
                
                # Keep only assignments not from this year
                remaining_assignments = assignments_data[assignments_data['tahun'] != year]
                success = storage_service.write_csv(remaining_assignments, 'config/checklist-assignments.csv')
                if success:
                    safe_print(f"✅ Deleted {deleted_summary['assignments']} assignments for year {year}")
        except Exception as e:
            safe_print(f"⚠️ Error deleting assignments: {e}")
        
        # 6. Delete uploaded files for the year
        try:
            files = storage_service.list_files(f"gcg-documents/{year}")
            if files:
                file_count = 0
                for file_info in files:
                    file_path = f"gcg-documents/{year}/{file_info['name']}"
                    success = storage_service.delete_file(file_path)
                    if success:
                        file_count += 1
                deleted_summary['uploaded_files'] = file_count
                safe_print(f"✅ Deleted {file_count} uploaded files for year {year}")
        except Exception as e:
            safe_print(f"⚠️ Error deleting uploaded files: {e}")
        
        # Calculate total deleted items
        total_deleted = (
            deleted_summary['checklist_items'] +
            deleted_summary['aspects'] +
            deleted_summary['users'] +
            deleted_summary['organizational_data']['direktorat'] +
            deleted_summary['organizational_data']['subdirektorat'] +
            deleted_summary['organizational_data']['divisi'] +
            deleted_summary['uploaded_files'] +
            deleted_summary['assignments']
        )
        
        safe_print(f"🎉 Bulk delete completed for year {year}. Total items deleted: {total_deleted}")
        
        return jsonify({
            'success': True,
            'message': f'Successfully deleted all data for year {year}',
            'year': year,
            'deleted_summary': deleted_summary,
            'total_deleted': total_deleted
        }), 200
        
    except Exception as e:
        safe_print(f"❌ Error during bulk delete: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/bulk-download-all-documents', methods=['POST'])
def bulk_download_all_documents():
    """Download all GCG and AOI documents organized by division, including checklist.csv"""
    import zipfile
    import io
    import tempfile
    from datetime import datetime

    try:
        data = request.get_json()
        year = data.get('year')
        include_gcg = data.get('includeGCG', True)
        include_aoi = data.get('includeAOI', True)
        include_checklist = data.get('includeChecklist', True)

        if not year:
            return jsonify({'error': 'Year is required'}), 400

        safe_print(f"🔍 Starting bulk download for year {year}")

        # Create a temporary file for the ZIP
        temp_zip = tempfile.NamedTemporaryFile(delete=False, suffix='.zip')
        temp_zip.close()  # Close the file handle immediately so zipfile can use it

        with zipfile.ZipFile(temp_zip.name, 'w', zipfile.ZIP_DEFLATED) as zipf:
            files_added = 0

            # Add checklist.csv for reference
            if include_checklist:
                try:
                    checklist_data = storage_service.read_csv(f'config/checklist-{year}.csv')
                    if checklist_data is None:
                        # Try without year suffix
                        checklist_data = storage_service.read_csv('config/checklist.csv')
                        if checklist_data is not None:
                            # Filter by year
                            checklist_data = checklist_data[checklist_data['tahun'] == year]

                    if checklist_data is not None and not checklist_data.empty:
                        csv_content = checklist_data.to_csv(index=False)
                        zipf.writestr(f'checklist_{year}.csv', csv_content)
                        files_added += 1
                        safe_print(f"✅ Added checklist_{year}.csv")
                except Exception as e:
                    safe_print(f"⚠️ Could not add checklist: {e}")

            # Download GCG documents organized by division from local storage
            if include_gcg:
                try:
                    safe_print(f"📄 Processing GCG documents...")
                    import glob

                    # Use local file system path
                    gcg_base_path = Path(__file__).parent.parent / 'data' / 'gcg-documents' / str(year)

                    if gcg_base_path.exists():
                        # Walk through all subdirectories and files
                        for file_path in gcg_base_path.rglob('*'):
                            if file_path.is_file():
                                try:
                                    # Get relative path from year folder
                                    rel_path = file_path.relative_to(gcg_base_path)

                                    # Read file content
                                    with open(file_path, 'rb') as f:
                                        file_content = f.read()

                                    # Add to ZIP with organized structure
                                    zip_path = f"GCG_Documents/{rel_path}"
                                    zipf.writestr(str(zip_path), file_content)
                                    files_added += 1
                                    safe_print(f"✅ Added GCG: {zip_path}")

                                except Exception as e:
                                    safe_print(f"❌ Failed to add GCG file {file_path}: {e}")
                    else:
                        safe_print(f"⚠️ GCG documents folder not found: {gcg_base_path}")

                except Exception as e:
                    safe_print(f"❌ Error processing GCG documents: {e}")

            # Download AOI documents from local storage (if exists)
            if include_aoi:
                try:
                    safe_print(f"📋 Processing AOI documents...")

                    # Use local file system path
                    aoi_base_path = Path(__file__).parent.parent / 'data' / 'aoi-documents' / str(year)

                    if aoi_base_path.exists():
                        # Walk through all subdirectories and files
                        for file_path in aoi_base_path.rglob('*'):
                            if file_path.is_file():
                                try:
                                    # Get relative path from year folder
                                    rel_path = file_path.relative_to(aoi_base_path)

                                    # Read file content
                                    with open(file_path, 'rb') as f:
                                        file_content = f.read()

                                    # Add to ZIP with organized structure
                                    zip_path = f"AOI_Documents/{rel_path}"
                                    zipf.writestr(str(zip_path), file_content)
                                    files_added += 1
                                    safe_print(f"✅ Added AOI: {zip_path}")

                                except Exception as e:
                                    safe_print(f"❌ Failed to add AOI file {file_path}: {e}")
                    else:
                        safe_print(f"⚠️ AOI documents folder not found: {aoi_base_path}")

                except Exception as e:
                    safe_print(f"❌ Error processing AOI documents: {e}")

            # Download Random documents (DOKUMEN_LAINNYA) from uploaded-files.xlsx
            try:
                safe_print(f"📂 Processing Random documents (DOKUMEN_LAINNYA)...")

                # Load uploaded files to find random documents
                files_data = storage_service.read_excel('uploaded-files.xlsx')
                if files_data is not None and not files_data.empty:
                    # Filter for random documents (checklistId is null/empty AND year matches)
                    random_docs = files_data[
                        (files_data['year'] == year) &
                        (files_data['checklistId'].isna() | (files_data['checklistId'] == ''))
                    ]

                    safe_print(f"📋 Found {len(random_docs)} random documents for year {year}")

                    for _, row in random_docs.iterrows():
                        try:
                            local_file_path = row.get('localFilePath', '')
                            if not local_file_path:
                                continue

                            # Construct full path
                            full_path = Path(__file__).parent.parent / 'data' / local_file_path

                            if full_path.exists():
                                # Read file content
                                with open(full_path, 'rb') as f:
                                    file_content = f.read()

                                # Add to ZIP with folder structure preserved
                                # Extract folder structure from localFilePath
                                # Format: random-documents/2024/folder_name/file.ext
                                zip_path = f"Dokumen_Lainnya/{local_file_path.replace('random-documents/' + str(year) + '/', '')}"
                                zipf.writestr(zip_path, file_content)
                                files_added += 1
                                safe_print(f"✅ Added Random: {zip_path}")
                            else:
                                safe_print(f"⚠️ Random document file not found: {full_path}")

                        except Exception as e:
                            safe_print(f"❌ Failed to add random document: {e}")
                else:
                    safe_print(f"⚠️ No uploaded-files.xlsx data found")

            except Exception as e:
                safe_print(f"❌ Error processing random documents: {e}")

        safe_print(f"📦 ZIP creation complete. Total files: {files_added}")

        if files_added == 0:
            return jsonify({'error': 'No documents found for the specified year'}), 404

        # Read the ZIP file and prepare response
        with open(temp_zip.name, 'rb') as zip_file:
            zip_data = zip_file.read()

        # Clean up temp file
        os.unlink(temp_zip.name)

        # Create response
        filename = f"All_Documents_{year}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        response = make_response(zip_data)
        response.headers['Content-Type'] = 'application/zip'
        response.headers['Content-Disposition'] = f'attachment; filename="{filename}"'
        response.headers['Content-Length'] = len(zip_data)

        safe_print(f"✅ Bulk download complete: {filename} ({len(zip_data)} bytes)")
        return response

    except Exception as e:
        safe_print(f"❌ Bulk download error: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'Failed to create bulk download: {str(e)}'}), 500

@app.route('/api/refresh-tracking-tables', methods=['POST'])
def refresh_tracking_tables():
    """
    Validate tracking files against actual storage and clean up orphaned records.
    Checks both uploaded-files.xlsx (GCG) and aoi-documents.csv (AOI) against actual files in storage.
    """
    try:
        data = request.get_json()
        year = data.get('year')

        if not year:
            return jsonify({'error': 'Year is required'}), 400

        safe_print(f"🔍 Starting tracking tables refresh for year {year}")

        gcg_cleaned = 0
        aoi_cleaned = 0

        # 1. Clean GCG documents tracking (uploaded-files.xlsx)
        try:
            safe_print(f"📄 Checking GCG documents tracking...")

            # Read uploaded-files.xlsx
            uploaded_files_data = storage_service.read_excel('uploaded-files.xlsx')
            if uploaded_files_data is not None and not uploaded_files_data.empty:
                initial_count = len(uploaded_files_data)

                # Filter for the specified year
                year_files = uploaded_files_data[uploaded_files_data['year'] == year].copy()
                other_years = uploaded_files_data[uploaded_files_data['year'] != year].copy()

                if not year_files.empty:
                    valid_records = []

                    for index, row in year_files.iterrows():
                        try:
                            # Extract file information
                            pic_name = str(row.get('subdirektorat', ''))
                            checklist_id = row.get('checklistId', '')
                            year_val = row.get('year', year)

                            if not pic_name or not checklist_id:
                                safe_print(f"⚠️ Skipping invalid record: missing PIC or checklist ID")
                                continue

                            # Clean PIC name for path
                            from werkzeug.utils import secure_filename
                            pic_clean = secure_filename(pic_name.replace(' ', '_'))

                            # Check if directory exists in local storage
                            directory_path = f"gcg-documents/{year_val}/{pic_clean}/{checklist_id}"
                            local_dir = Path(__file__).parent.parent / 'data' / directory_path

                            try:
                                # Check if directory has actual files (not just placeholders)
                                has_real_files = False
                                if local_dir.exists() and local_dir.is_dir():
                                    for file_item in local_dir.iterdir():
                                        if (file_item.is_file() and
                                            file_item.name != '.emptyFolderPlaceholder' and
                                            not file_item.name.startswith('.')):
                                            has_real_files = True
                                            break

                                if has_real_files:
                                    valid_records.append(row)
                                    safe_print(f"✅ Valid GCG record: {directory_path}")
                                else:
                                    safe_print(f"❌ Orphaned GCG record (no files): {directory_path}")
                                    gcg_cleaned += 1

                            except Exception as e:
                                safe_print(f"❌ Error checking GCG directory {directory_path}: {e}")
                                gcg_cleaned += 1

                        except Exception as e:
                            safe_print(f"❌ Error processing GCG record: {e}")
                            gcg_cleaned += 1

                    # Rebuild the dataframe with valid records + other years
                    if valid_records:
                        valid_year_df = pd.DataFrame(valid_records)
                        updated_df = pd.concat([other_years, valid_year_df], ignore_index=True)
                    else:
                        updated_df = other_years

                    # Save cleaned data
                    if len(updated_df) != initial_count:
                        success = storage_service.write_excel(updated_df, 'uploaded-files.xlsx')
                        if success:
                            safe_print(f"✅ Cleaned {gcg_cleaned} orphaned GCG records")
                        else:
                            safe_print(f"❌ Failed to save cleaned GCG data")

        except Exception as e:
            safe_print(f"❌ Error cleaning GCG tracking: {e}")

        # 2. Clean AOI documents tracking (aoi-documents.csv)
        try:
            safe_print(f"📋 Checking AOI documents tracking...")

            # Read aoi-documents.csv
            aoi_docs_data = storage_service.read_csv('config/aoi-documents.csv')
            if aoi_docs_data is not None and not aoi_docs_data.empty:
                initial_count = len(aoi_docs_data)

                # Filter for the specified year
                year_docs = aoi_docs_data[aoi_docs_data['tahun'] == year].copy()
                other_years = aoi_docs_data[aoi_docs_data['tahun'] != year].copy()

                if not year_docs.empty:
                    valid_records = []

                    for index, row in year_docs.iterrows():
                        try:
                            # Extract file information
                            subdirektorat = str(row.get('subdirektorat', ''))
                            recommendation_id = row.get('aoiRecommendationId', '')
                            year_val = row.get('tahun', year)

                            if not subdirektorat or not recommendation_id:
                                safe_print(f"⚠️ Skipping invalid AOI record: missing subdirektorat or recommendation ID")
                                continue

                            # Clean subdirektorat name for path
                            from werkzeug.utils import secure_filename
                            subdirektorat_clean = secure_filename(subdirektorat.replace(' ', '_'))

                            # Check if directory exists in local storage
                            directory_path = f"aoi-documents/{year_val}/{subdirektorat_clean}/{recommendation_id}"
                            local_dir = Path(__file__).parent.parent / 'data' / directory_path

                            try:
                                # Check if directory has actual files (not just placeholders)
                                has_real_files = False
                                if local_dir.exists() and local_dir.is_dir():
                                    for file_item in local_dir.iterdir():
                                        if (file_item.is_file() and
                                            file_item.name != '.emptyFolderPlaceholder' and
                                            not file_item.name.startswith('.')):
                                            has_real_files = True
                                            break

                                if has_real_files:
                                    valid_records.append(row)
                                    safe_print(f"✅ Valid AOI record: {directory_path}")
                                else:
                                    safe_print(f"❌ Orphaned AOI record (no files): {directory_path}")
                                    aoi_cleaned += 1

                            except Exception as e:
                                safe_print(f"❌ Error checking AOI directory {directory_path}: {e}")
                                aoi_cleaned += 1

                        except Exception as e:
                            safe_print(f"❌ Error processing AOI record: {e}")
                            aoi_cleaned += 1

                    # Rebuild the dataframe with valid records + other years
                    if valid_records:
                        valid_year_df = pd.DataFrame(valid_records)
                        updated_df = pd.concat([other_years, valid_year_df], ignore_index=True)
                    else:
                        updated_df = other_years

                    # Save cleaned data
                    if len(updated_df) != initial_count:
                        success = storage_service.write_csv(updated_df, 'config/aoi-documents.csv')
                        if success:
                            safe_print(f"✅ Cleaned {aoi_cleaned} orphaned AOI records")
                        else:
                            safe_print(f"❌ Failed to save cleaned AOI data")

        except Exception as e:
            safe_print(f"❌ Error cleaning AOI tracking: {e}")

        safe_print(f"🎉 Tracking tables refresh complete for year {year}")
        safe_print(f"📊 Summary: GCG cleaned: {gcg_cleaned}, AOI cleaned: {aoi_cleaned}")

        return jsonify({
            'success': True,
            'message': f'Tracking tables refreshed for year {year}',
            'year': year,
            'gcgCleaned': gcg_cleaned,
            'aoiCleaned': aoi_cleaned,
            'totalCleaned': gcg_cleaned + aoi_cleaned
        }), 200

    except Exception as e:
        safe_print(f"❌ Error during tracking tables refresh: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'Failed to refresh tracking tables: {str(e)}'}), 500


if __name__ == '__main__':
    import os
    import socket

    # Try to find an available port starting from the default
    default_port = int(os.environ.get('FLASK_PORT', 5001))
    port = default_port
    max_attempts = 10

    for attempt in range(max_attempts):
        try:
            # Test if port is available
            test_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            test_socket.bind(('0.0.0.0', port))
            test_socket.close()
            break  # Port is available
        except OSError:
            if attempt < max_attempts - 1:
                port += 1
            else:
                safe_print(f"❌ Could not find available port after {max_attempts} attempts starting from {default_port}")
                sys.exit(1)

    if port != default_port:
        safe_print(f"⚠️  WARNING: Default port {default_port} was busy")
        safe_print(f"⚠️  Backend running on port {port} instead")
        safe_print(f"⚠️  Update vite.config.ts proxy target to: http://localhost:{port}")

    app.run(debug=True, host='0.0.0.0', port=port, use_reloader=False)
