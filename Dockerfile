# Use a Python base image
FROM python:3.9-slim

# Set the working directory in the container
WORKDIR /app

# Copy requirements.txt into the container
COPY requirements.txt /app/

# Upgrade pip and install dependencies
RUN pip install --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# Install system dependencies for tkinter, fonts, and Xvfb
RUN apt-get update && \
    apt-get install -y \
        python3-tk \
        ttf-wqy-zenhei \
        libxrender1 \
        libxext6 \
        libx11-6 \
        xvfb && \
    apt-get clean && rm -rf /var/lib/apt/lists/*

# Set environment variables for Xvfb
ENV DISPLAY=:99

# Copy the rest of the application files into the container
COPY . /app

# Expose the port if needed (optional)
EXPOSE 5000

# Start Xvfb and then run the application
CMD ["sh", "-c", "Xvfb :99 -screen 0 1024x768x16 & python Word_2_Pdf.py"]
