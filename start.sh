#!/bin/bash
# GCG Document Hub - Linux/Mac Start Script
# This script checks dependencies and starts the development server

set -e

echo "=========================================="
echo "  GCG Document Hub - Starting Server"
echo "=========================================="
echo ""

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Check if Node.js is installed
if ! command -v node &> /dev/null; then
    echo -e "${RED}❌ Node.js is not installed!${NC}"
    echo "Please install Node.js from: https://nodejs.org/"
    exit 1
fi

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo -e "${RED}❌ Python 3 is not installed!${NC}"
    echo "Please install Python 3 from: https://www.python.org/"
    exit 1
fi

echo -e "${GREEN}✅ Node.js $(node --version) detected${NC}"
echo -e "${GREEN}✅ Python $(python3 --version | awk '{print $2}') detected${NC}"
echo ""

# Check if node_modules exists
if [ ! -d "node_modules" ]; then
    echo -e "${YELLOW}📦 Installing Node.js dependencies...${NC}"
    npm install
    echo -e "${GREEN}✅ Node.js dependencies installed${NC}"
    echo ""
else
    echo -e "${GREEN}✅ Node.js dependencies found${NC}"
fi

# Check if Python virtual environment exists
if [ ! -d "venv" ]; then
    echo -e "${YELLOW}🐍 Creating Python virtual environment...${NC}"
    python3 -m venv venv
    echo -e "${GREEN}✅ Virtual environment created${NC}"
    echo ""
fi

# Activate virtual environment and install dependencies
echo -e "${YELLOW}🐍 Installing Python dependencies...${NC}"
source venv/bin/activate
pip install -q --upgrade pip
pip install -q -r backend/requirements.txt
echo -e "${GREEN}✅ Python dependencies installed${NC}"
echo ""

# Check if .env file exists
if [ ! -f ".env" ]; then
    echo -e "${YELLOW}⚠️  .env file not found${NC}"
    echo "Creating .env from .env.example (if exists)..."
    if [ -f ".env.example" ]; then
        cp .env.example .env
        echo -e "${GREEN}✅ .env file created - please configure it${NC}"
    else
        echo -e "${RED}❌ No .env.example found - you may need to create .env manually${NC}"
    fi
    echo ""
fi

echo "=========================================="
echo -e "${GREEN}🚀 Starting Development Server...${NC}"
echo "=========================================="
echo ""
echo "Frontend: http://localhost:8080"
echo "Backend:  http://localhost:5001"
echo ""
echo "Press Ctrl+C to stop the server"
echo ""

# Start the development server
npm run dev
