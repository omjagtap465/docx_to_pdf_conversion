Before running this project, ensure you have the following installed:

1. .NET Core SDK (version 6.0 or later)
2. LibreOffice (for .docx to .pdf conversion)
3. Spire.Pdf package
Download Libre Office from https://www.libreoffice.org/download/download-libreoffice/

Setup Instructions
1. Clone the repository
   git clone https://github.com/omjagtap465/docx_to_pdf_conversion.git

2. Install Dependencies
   dotnet restore
3. Install Required Packages
   dotnet add package Spire.Pdf
   dotnet add package DocXToPdfConverter --version 1.0.5

4. Set Up LibreOffice Path
   Ensure that LibreOffice is installed and update the locationOfLibreOfficeSoffice variable in Program.cs with the correct path.

5. Build and Run

  To build the project, run:

  dotnet build

  To execute the converter, use:

  dotnet run <directory-path>

  Replace <directory-path> with the folder path 
  Example :
  dotnet run "C:\Users\lenovo\Desktop\wordFiles\"

   
