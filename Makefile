# Makefile for Advanced Finance Dashboard

.PHONY: help install install-dev test test-coverage lint format clean run build docker-build docker-run

# Default target
help:
	@echo "Advanced Finance Dashboard - Available Commands:"
	@echo ""
	@echo "  install      - Install production dependencies"
	@echo "  install-dev  - Install development dependencies" 
	@echo "  test         - Run all tests"
	@echo "  test-coverage - Run tests with coverage report"
	@echo "  test-100     - Run 100 consecutive test runs"
	@echo "  lint         - Run code linting"
	@echo "  format       - Format code with black and isort"
	@echo "  clean        - Clean temporary files"
	@echo "  run          - Run the dashboard application"
	@echo "  build        - Build distribution packages"
	@echo "  docker-build - Build Docker image"
	@echo "  docker-run   - Run Docker container"

# Installation
install:
	pip install -r requirements.txt

install-dev:
	pip install -r requirements.txt
	pip install pytest pytest-cov pytest-mock black flake8 mypy isort coverage

# Testing
test:
	python -m pytest tests/ -v

test-coverage:
	python -m pytest tests/ --cov=src --cov-report=html --cov-report=term-missing

test-100:
	python run_100_consecutive.py

# Code Quality
lint:
	flake8 src tests
	mypy src tests

format:
	black src tests
	isort src tests

# Maintenance
clean:
	find . -type d -name "__pycache__" -exec rm -rf {} +
	find . -type f -name "*.pyc" -delete
	find . -type f -name "*.pyo" -delete
	find . -type d -name "*.egg-info" -exec rm -rf {} +
	rm -rf build/ dist/ .coverage htmlcov/ .pytest_cache/

# Application
run:
	python main_dashboard.py

# Build
build: clean
	python setup.py sdist bdist_wheel

# Docker
docker-build:
	docker build -t finance-dashboard .

docker-run:
	docker run -p 8080:8080 -p 8081:8081 finance-dashboard