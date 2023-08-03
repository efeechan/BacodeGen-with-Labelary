import os
import sys
import requests
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

# Add this page size dictionary
page_sizes = {
    "A4": (297, 210),
    "A5": (210, 148),
    "A6": (148, 105),
    "Letter": (215.9, 279.4),
    "Legal": (215.9, 355.6),
}

# Add landscape versions of the page sizes
landscape_page_sizes = {
    size: (width, height) for size, (height, width) in page_sizes.items()
}

# Define page_size_var, page_length_var, and page_width_var as global variables
page_size_var = None
page_length_var = None
page_width_var = None
page_orientation_var = None

def process_files(root, input_dir, output_dir, file_format, page_length, page_width):
    # Check if the specified file format is valid
    if file_format not in ["pdf", "png"]:
        messagebox.showerror(
            "Invalid File Format", "Please specify either 'pdf' or 'png'."
        )
        return

    # Check if the specified page length and width are within the valid range
    if page_length == 0 or page_length < 0 or page_length > 381:
        messagebox.showerror(
            "Invalid Page Length",
            "Please specify a value between 0 and 381 millimeters for page length.",
        )
        return

    if page_width == 0 or page_width < 0 or page_width > 381:
        messagebox.showerror(
            "Invalid Page Width",
            "Please specify a value between 0 and 381 millimeters for page width.",
        )
        return

    # Create a scrolled text widget to show the print messages
    text_widget = scrolledtext.ScrolledText(
        root, wrap=tk.WORD, width=40, height=10, state=tk.DISABLED
    )
    text_widget.grid(row=7, column=0, columnspan=4, padx=5, pady=5)

    # Redirect print to the text widget
    def print_to_text_widget(message):
        text_widget.configure(state=tk.NORMAL)
        text_widget.insert(tk.END, message)
        text_widget.see(tk.END)
        text_widget.configure(state=tk.DISABLED)
        root.update()

    sys.stdout.write = print_to_text_widget
    sys.stderr.write = print_to_text_widget

    # Iterate over the files in the input directory
    for filename in os.listdir(input_dir):
        if filename.endswith(".zpl"):
            input_file_path = os.path.join(input_dir, filename)

            # Read the content of the ZPL file
            with open(input_file_path, "r") as file:
                zpl_content = file.read()

            # Adjust the Labelary API URL and desired output format as needed
            url = f"http://api.labelary.com/v1/printers/8dpmm/labels/{page_width / 25.4}x{page_length / 25.4}/0/"
            files = {"file": ("label.zpl", zpl_content)}

            if file_format == "pdf":
                print(f"Processing file {filename} for PDF")
                # Generate PDF file
                pdf_headers = {"Accept": "application/pdf"}
                pdf_response = requests.post(
                    url, headers=pdf_headers, files=files, stream=True
                )

                if pdf_response.status_code == 200:
                    pdf_response.raw.decode_content = True

                    # Prepare the output file path and name for PDF
                    output_pdf_path = os.path.join(
                        output_dir, os.path.splitext(filename)[0] + ".pdf"
                    )

                    # Save the PDF response content to the output file
                    with open(output_pdf_path, "wb") as pdf_out_file:
                        shutil.copyfileobj(pdf_response.raw, pdf_out_file)
                else:
                    print(
                        f"Error processing file {filename} for PDF: {pdf_response.text}"
                    )

            if file_format == "png":
                print(f"Processing file {filename} for PNG")
                # Generate PNG file
                png_headers = {"Accept": "image/png"}
                png_response = requests.post(
                    url, headers=png_headers, files=files, stream=True
                )

                if png_response.status_code == 200:
                    png_response.raw.decode_content = True

                    # Prepare the output file path and name for PNG
                    output_png_path = os.path.join(
                        output_dir, os.path.splitext(filename)[0] + ".png"
                    )

                    # Save the PNG response content to the output file
                    with open(output_png_path, "wb") as png_out_file:
                        shutil.copyfileobj(png_response.raw, png_out_file)
                else:
                    print(
                        f"Error processing file {filename} for PNG: {png_response.text}"
                    )
        else:
            print(f"Skipped file {filename}: Not a ZPL file")

    # Print a message when processing is complete
    print("Conversion completed successfully!")


def main():
    root = tk.Tk()
    root.title("ZPL to PDF/PNG Converter")

    # Create Tkinter variables to store user input
    input_dir_var = tk.StringVar()
    output_dir_var = tk.StringVar()
    file_format_var = tk.StringVar(value="pdf")  # Default value is "pdf"
    page_length_var = tk.DoubleVar(value=100.0)  # Default value is 100.0
    page_width_var = tk.DoubleVar(value=150.0)  # Default value is 150.0

    # Function to handle the page orientation change
    def update_page_orientation(selected_orientation):

        # Get the selected page size from the dropdown
        selected_page_size = page_size_var.get()
        
        if selected_page_size != "Custom":
            # Set the correct landscape page size based on the selected orientation
            if selected_orientation == "Landscape":
                page_length, page_width = landscape_page_sizes[selected_page_size]
            else:
                page_length, page_width = page_sizes[selected_page_size]

            # Update the entry fields with the correct values based on the orientation
            page_length_var.set(page_length)
            page_width_var.set(page_width)

    def update_page_size_option(selected_option):
        # Show/hide page length and width entry fields based on the selected option
        custom_selected = selected_option == "Custom"
        page_length_entry.config(state=tk.NORMAL if custom_selected else tk.DISABLED)
        page_width_entry.config(state=tk.NORMAL if custom_selected else tk.DISABLED)

        if custom_selected:
            # Disable page orientation menu when "Custom" is selected
            page_orientation_menu.configure(state=tk.DISABLED)
        else:
            # Enable page orientation menu when a predefined size is selected
            page_orientation_menu.configure(state=tk.NORMAL)

        if not custom_selected:
            # Set the correct landscape page size based on the selected orientation
            if page_orientation_var.get() == "Landscape":
                page_length, page_width = landscape_page_sizes[selected_option]
            else:
                page_length, page_width = page_sizes[selected_option]

            # Update the entry fields with the correct values based on the orientation
            page_length_var.set(page_length)
            page_width_var.set(page_width)


    # Function to handle the Convert button click event
    def convert_files():
        input_dir = input_dir_var.get()
        output_dir = output_dir_var.get()
        file_format = file_format_var.get()

        # Get the selected page size from the dropdown
        selected_page_size = page_size_var.get()

        # Get the selected page orientation
        selected_orientation = page_orientation_var.get()

        # Set the correct landscape page size based on the selected orientation
        if selected_orientation == "Landscape":
            if selected_page_size == "Custom":
                page_length = page_width_var.get()
                page_width = page_length_var.get()
            else:
                page_length, page_width = landscape_page_sizes[selected_page_size]
        else:
            if selected_page_size == "Custom":
                page_length = page_length_var.get()
                page_width = page_width_var.get()
            else:
                page_length, page_width = page_sizes[selected_page_size]

        # Adjust the Labelary API URL and desired output format with the updated page length and width
        url = f"http://api.labelary.com/v1/printers/8dpmm/labels/{page_width / 25.4}x{page_length / 25.4}/0/"

        # Continue with the rest of your existing code for processing files
        process_files(
            root,
            input_dir,
            output_dir,
            file_format,
            page_length,
            page_width,
        )

    # Function to browse and set the input directory path
    def browse_input_dir():
        input_dir = filedialog.askdirectory()
        if input_dir:
            input_dir_var.set(input_dir)

    # Function to browse and set the output directory path
    def browse_output_dir():
        output_dir = filedialog.askdirectory()
        if output_dir:
            output_dir_var.set(output_dir)

    # Create the GUI layout
    tk.Label(root, text="Input Directory:").grid(row=0, column=0, padx=5, pady=5)
    tk.Entry(root, textvariable=input_dir_var).grid(
        row=0, column=1, padx=5, pady=5, columnspan=2
    )
    tk.Button(root, text="Browse", command=browse_input_dir).grid(
        row=0, column=3, padx=5, pady=5
    )

    tk.Label(root, text="Output Directory:").grid(row=1, column=0, padx=5, pady=5)
    tk.Entry(root, textvariable=output_dir_var).grid(
        row=1, column=1, padx=5, pady=5, columnspan=2
    )
    tk.Button(root, text="Browse", command=browse_output_dir).grid(
        row=1, column=3, padx=5, pady=5
    )

    tk.Label(root, text="File Format:").grid(row=2, column=0, padx=5, pady=5)
    tk.Radiobutton(root, text="PDF", variable=file_format_var, value="pdf").grid(
        row=2, column=1, padx=5, pady=5
    )
    tk.Radiobutton(root, text="PNG", variable=file_format_var, value="png").grid(
        row=2, column=2, padx=5, pady=5
    )

    # New options for page size
    page_size_options = ["Custom", "A4", "A5", "A6", "Letter", "Legal"]
    page_size_var = tk.StringVar(
        value=page_size_options[0]
    )  # Default value is "Custom"
    tk.Label(root, text="Page Size:").grid(row=3, column=0, padx=5, pady=5)
    page_size_menu = tk.OptionMenu(
        root, page_size_var, *page_size_options, command=update_page_size_option
    )
    page_size_menu.grid(row=3, column=1, columnspan=2, padx=5, pady=5)

    # New input fields for custom page length and width
    tk.Label(root, text="Page Length (mm):").grid(row=4, column=0, padx=5, pady=5)
    page_length_entry = tk.Entry(root, textvariable=page_length_var)
    page_length_entry.grid(
        row=4, column=1, padx=5, pady=5, columnspan=3
    )  # Use columnspan=3 to span over multiple columns

    tk.Label(root, text="Page Width (mm):").grid(row=5, column=0, padx=5, pady=5)
    page_width_entry = tk.Entry(root, textvariable=page_width_var)
    page_width_entry.grid(
        row=5, column=1, padx=5, pady=5, columnspan=3
    )  # Use columnspan=3 to span over multiple columns

    page_orientation_options = ["Portrait", "Landscape"]
    page_orientation_var = tk.StringVar(value=page_orientation_options[0])  # Default value is "Portrait"
    tk.Label(root, text="Page Orientation:").grid(row=6, column=0, padx=5, pady=5)
    page_orientation_menu = tk.OptionMenu(root, page_orientation_var, *page_orientation_options, command=update_page_orientation)
    page_orientation_menu.grid(row=6, column=1, columnspan=2, padx=5, pady=5)

    tk.Button(root, text="Convert", command=convert_files).grid(
        row=8, column=1, columnspan=1, padx=5, pady=5
    )

    # Set the default page size to "Custom" and disable the page orientation menu
    page_size_var.set(page_size_options[0])
    update_page_size_option(page_size_options[0])

    root.mainloop()


if __name__ == "__main__":
    main()
