# Use an official Python runtime as a parent image
FROM python:3.10-slim

# Set the working directory in the container
WORKDIR /app

# Copy the dependencies file to the working directory
COPY requirements.txt .

# Install any needed packages specified in requirements.txt
# Dependencies are managed in requirements.txt to avoid conflicts.
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of the application code to the working directory
COPY . .

# Set the environment variable for the port.
# This value is provided by Cloud Run at runtime. Default is 8080.
ENV PORT 8080

# Command to run the application.
# It uses the $PORT environment variable to listen on the correct port.
CMD uvicorn main:app --host 0.0.0.0 --port $PORT
