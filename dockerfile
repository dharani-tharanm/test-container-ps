# Use the latest PowerShell image
FROM mcr.microsoft.com/powershell:lts-alpine-3.13

# Set the working directory in the container
WORKDIR /usr/src/app

# Copy your PowerShell script into the container
COPY ./main .

COPY ./Store.xlsx .
COPY ./User.xlsx .

# Set the PowerShell script to run when the container starts
CMD ["pwsh", "-File", "./main.ps1"]
