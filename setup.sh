#!/bin/bash

echo "========================================="
echo "NUS Canvas Deadline Scraper - Setup"
echo "========================================="
echo ""

# Check Python version
python_version=$(python3 --version 2>&1 | awk '{print $2}')
echo "✓ Found Python $python_version"

# Install dependencies
echo ""
echo "Installing dependencies..."
pip install -r requirements.txt

# Create .env if it doesn't exist
if [ ! -f .env ]; then
    echo ""
    echo "Creating .env file..."
    cp .env.example .env
    echo "✓ Created .env file"
    echo ""
    echo "⚠️  IMPORTANT: Edit .env and add your API tokens!"
    echo ""
    echo "You need:"
    echo "1. Canvas API token from https://canvas.nus.edu.sg/profile/settings"
    echo "2. Anthropic API key from https://console.anthropic.com/"
    echo ""
    echo "Run this to edit:"
    echo "  nano .env"
    echo ""
else
    echo ""
    echo "✓ .env file already exists"
fi

# Create directories
mkdir -p downloaded_materials
echo "✓ Created downloaded_materials directory"

echo ""
echo "========================================="
echo "Setup complete!"
echo "========================================="
echo ""
echo "Next steps:"
echo "1. Edit .env and add your API tokens"
echo "2. Run: python3 scrape_deadlines.py"
echo ""
echo "See QUICKSTART.md for detailed instructions"
echo ""
