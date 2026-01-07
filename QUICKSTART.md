# GCG Document Hub - Quick Start Guide

## Prerequisites

Before running the application, make sure you have:

- **Node.js** (v16 or higher) - [Download](https://nodejs.org/)
- **Python** (v3.8 or higher) - [Download](https://www.python.org/)
- **Git** (optional, for cloning) - [Download](https://git-scm.com/)

## Quick Start (Easiest Way)

### For Linux/Mac Users:

```bash
./start.sh
```

### For Windows Users:

```cmd
start.bat
```

**That's it!** The script will:
- ✅ Check if Node.js and Python are installed
- ✅ Install all Node.js dependencies (npm packages)
- ✅ Create Python virtual environment
- ✅ Install all Python dependencies
- ✅ Start both frontend and backend servers

## What the Start Script Does

1. **Dependency Check**: Verifies Node.js and Python are installed
2. **Node.js Setup**: Installs packages from `package.json` if needed
3. **Python Setup**: Creates virtual environment and installs from `requirements.txt`
4. **Environment File**: Creates `.env` from `.env.example` if missing
5. **Server Start**: Launches both frontend (port 8080) and backend (port 5001)

## Manual Installation (If Scripts Don't Work)

### Step 1: Install Node.js Dependencies

```bash
npm install
```

### Step 2: Install Python Dependencies

**Linux/Mac:**
```bash
python3 -m venv venv
source venv/bin/activate
pip install -r backend/requirements.txt
```

**Windows:**
```cmd
python -m venv venv
venv\Scripts\activate.bat
pip install -r backend\requirements.txt
```

### Step 3: Configure Environment

```bash
# Copy example environment file (if exists)
cp .env.example .env

# Edit .env file with your Supabase credentials
```

### Step 4: Start Development Server

```bash
npm run dev
```

## Accessing the Application

Once started, the application will be available at:

- **Frontend**: http://localhost:8080
- **Backend API**: http://localhost:5001

## Troubleshooting

### Port Already in Use

If you see "port already in use" error:

**Backend (automatic)**: The backend will automatically try ports 5001 → 5002 → 5003, etc.
- If it uses a different port, update `vite.config.ts` proxy target accordingly

**Frontend**: Kill the process using port 8080:
```bash
# Linux/Mac
lsof -ti:8080 | xargs kill -9

# Windows
netstat -ano | findstr :8080
taskkill /PID <PID> /F
```

### Python Virtual Environment Issues

**Linux/Mac:**
```bash
rm -rf venv
python3 -m venv venv
source venv/bin/activate
pip install -r backend/requirements.txt
```

**Windows:**
```cmd
rmdir /s /q venv
python -m venv venv
venv\Scripts\activate.bat
pip install -r backend\requirements.txt
```

### Node Modules Issues

```bash
rm -rf node_modules package-lock.json
npm install
```

## Development Commands

```bash
# Start development server (both frontend and backend)
npm run dev

# Start only frontend
npm run vite

# Start only backend
python backend/app.py

# Install new dependency
npm install <package-name>
pip install <package-name>
```

## Project Structure

```
pos-gcg-document-hub-v2/
├── backend/              # Flask backend
│   ├── app.py           # Main backend server
│   └── requirements.txt # Python dependencies
├── src/                 # React frontend
├── public/              # Static assets
├── start.sh            # Linux/Mac start script
├── start.bat           # Windows start script
└── package.json        # Node.js dependencies
```

## Need Help?

If you encounter any issues:
1. Check the error message in the terminal
2. Make sure all prerequisites are installed
3. Try the manual installation steps
4. Check if ports 8080 and 5001 are available

---

**Happy coding!** 🚀
