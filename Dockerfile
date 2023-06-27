# Use an official Python runtime as the base image
FROM python:3.9-slim

# Set the working directory in the container
WORKDIR /app

# Copy the requirements file
COPY requirements.txt .

# Install the Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy the script and data files to the container
COPY frotas_viagens.py .
COPY demanda_carreira.xlsx .
COPY relatorio.xlsx .

# Run the Python script when the container launches
CMD ["python", "frotas_viagens.py"]