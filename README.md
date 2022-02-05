# Abstrak
Easily generate focused resumes with consistent styles

## Requirements
- Python=>3.9.9
- PyYAML==6.0
- lxml==4.7.1
- python-docx==0.8.11
- jmespath==0.10.0

## Getting Started
```sh
# Clone the repo and enter the directory
git clone https://github.com/dylanturn/abstrak.git
cd abstrak

#Create the virtual environment
python3.9 -m virtualenv venv

# source the virtual environment
source venv/bin/activate

# Install the requirements
pip install -r requirements.txt

# Generate a resume!
python generate.py example-resume-data.yml
```

## TODO:
- Decouple styles?