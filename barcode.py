import os
import sys
import requests
import shutil
from gooey import Gooey

def main():
    # Get input and output file paths from command line arguments
    input_dir = sys.argv[1]  # Directory path where input ZPL files are located
    output_dir = sys.argv[2]  # Directory path where output files will be saved
    file_format = sys.argv[3]  # Output file format ('pdf' or 'png')

    # Check if the specified file format is valid
    if file_format not in ['pdf', 'png']:
        print('Invalid file format. Please specify either "pdf" or "png".')
        sys.exit()

    # Iterate over the files in the input directory
    for filename in os.listdir(input_dir):
        if filename.endswith('.zpl'):
            input_file_path = os.path.join(input_dir, filename)

            # Read the content of the ZPL file
            with open(input_file_path, 'r') as file:
                zpl_content = file.read()

            # Adjust the Labelary API URL and desired output format as needed
            url = 'http://api.labelary.com/v1/printers/8dpmm/labels/15x15/0/'
            files = {'file': ('label.zpl', zpl_content)}

            if file_format == 'pdf':
                # Generate PDF file
                pdf_headers = {'Accept': 'application/pdf'}
                pdf_response = requests.post(url, headers=pdf_headers, files=files, stream=True)

                if pdf_response.status_code == 200:
                    pdf_response.raw.decode_content = True

                    # Prepare the output file path and name for PDF
                    output_pdf_path = os.path.join(output_dir, os.path.splitext(filename)[0] + '.pdf')

                    # Save the PDF response content to the output file
                    with open(output_pdf_path, 'wb') as pdf_out_file:
                        shutil.copyfileobj(pdf_response.raw, pdf_out_file)
                else:
                    print(f'Error processing file {filename} for PDF: {pdf_response.text}')

            if file_format == 'png':
                # Generate PNG file
                png_headers = {'Accept': 'image/png'}
                png_response = requests.post(url, headers=png_headers, files=files, stream=True)

                if png_response.status_code == 200:
                    png_response.raw.decode_content = True

                    # Prepare the output file path and name for PNG
                    output_png_path = os.path.join(output_dir, os.path.splitext(filename)[0] + '.png')

                    # Save the PNG response content to the output file
                    with open(output_png_path, 'wb') as png_out_file:
                        shutil.copyfileobj(png_response.raw, png_out_file)
                else:
                    print(f'Error processing file {filename} for PNG: {png_response.text}')
        else:
            print(f'Skipped file {filename}: Not a ZPL file')

if __name__ == '__main__':
    main()