import os
import sys
import requests
import shutil
from gooey import Gooey, GooeyParser

@Gooey(program_name='ZPL to PDF/PNG Converter', default_size=(800, 600))
def main():
    # Create the GooeyParser
    parser = GooeyParser(description='Convert ZPL files to PDF or PNG')

    # Add the input directory argument
    parser.add_argument('input_dir', metavar='Input Directory', widget='DirChooser',
                        help='Directory path where input ZPL files are located')

    # Add the output directory argument
    parser.add_argument('output_dir', metavar='Output Directory', widget='DirChooser',
                        help='Directory path where output files will be saved')

    # Add the file format argument
    parser.add_argument('file_format', metavar='File Format', choices=['pdf', 'png'],
                        help='Output file format ("pdf" or "png")')

    # Add the page length argument
    parser.add_argument('--width', metavar='Page Length', type=float, default=15,
                        help='Width of the output page in inches (default: 15)')

    # Add the page width argument
    parser.add_argument('--length', metavar='Page Width', type=float, default=15,
                        help='Length of the output page in inches (default: 15)')

    # Parse the command-line arguments
    args = parser.parse_args()

    # Extract the input and output file paths, format, length, and width from the parsed arguments
    input_dir = args.input_dir
    output_dir = args.output_dir
    file_format = args.file_format
    page_length = args.length
    page_width = args.width

    # Check if the specified file format is valid
    if file_format not in ['pdf', 'png']:
        print('Invalid file format. Please specify either "pdf" or "png".')
        sys.exit()

    # Check if the specified page length and width are within the valid range
    if page_length <= 0 or page_width <= 0 or page_length > 15 or page_width > 15:
        print('Invalid page length or width. Please specify values between 0 and 15 inches.')
        sys.exit()

    # Iterate over the files in the input directory
    for filename in os.listdir(input_dir):
        if filename.endswith('.zpl'):
            input_file_path = os.path.join(input_dir, filename)

            # Read the content of the ZPL file
            with open(input_file_path, 'r') as file:
                zpl_content = file.read()

            # Adjust the Labelary API URL and desired output format as needed
            url = f'http://api.labelary.com/v1/printers/8dpmm/labels/{page_length}x{page_width}/0/'
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
