# Makefile for Mail Merge App

# Name of the main Python file (entry point)
MAIN = main.py

# Create virtual environment folder
VENV = venv

# Install dependencies into venv
.PHONY: install
install:
	python -m venv $(VENV)
	$(VENV)/Scripts/pip install --upgrade pip
	$(VENV)/Scripts/pip install pandas jinja2 smtplib tk pyinstaller

# Run the app
.PHONY: run
run:
	python $(MAIN)

# Build executable
.PHONY: build
build:
	pyinstaller --onefile --windowed $(MAIN)

# Clean build artifacts
.PHONY: clean
clean:
	rd /s /q build dist __pycache__ || true
	del /f /q *.spec || true
