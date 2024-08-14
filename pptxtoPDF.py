import os
import comtypes.client  
import argparse


def convert(input_path, output_folder_path):
    if os.path.isdir(input_path):
        # Input path is a directory
        input_folder_path = os.path.abspath(input_path)

        if not os.path.isdir(input_folder_path):
            print("Error: Input folder does not exist.")
            return

        input_file_paths = [os.path.join(input_folder_path, file_name) for file_name in os.listdir(input_folder_path)]
    else:
        # Input path is a file
        input_file_paths = [os.path.abspath(input_path)]

    # Use the input_file_path's directory if output_folder_path is not provided
    if not output_folder_path:
        output_folder_path = os.path.dirname(input_file_paths[0])

    output_folder_path = os.path.abspath(output_folder_path)

    if not os.path.exists(output_folder_path):
        os.makedirs(output_folder_path)

    success_count = 0
    error_count = 0

    for input_file_path in input_file_paths:
        if not input_file_path.lower().endswith((".ppt", ".pptx")):
            print(f"Skipping file '{input_file_path}' as it does not have a PowerPoint extension.")
            continue

        try:
            # Create PowerPoint application object
            powerpoint = comtypes.client.CreateObject("Powerpoint.Application")

            # Open the PowerPoint slides
            slides = powerpoint.Presentations.Open(input_file_path, WithWindow=False)

            # Get base file name
            file_name = os.path.splitext(os.path.basename(input_file_path))[0]

            # Create output file path
            output_file_path = os.path.join(output_folder_path, file_name + ".pdf")

            if os.path.exists(output_file_path):
                print(f"Error: Output file '{output_file_path}' already exists.")
                error_count += 1
                continue

            # Save as PDF (formatType = 32)
            slides.SaveAs(output_file_path, 32)

            # Close the slide deck
            slides.Close()

            powerpoint.Quit()

            success_count += 1
        except Exception as e:
            print(f"Error converting file '{input_file_path}': {str(e)}")
            error_count += 1

    print(f"Conversion completed: {success_count} files converted successfully, {error_count} files failed.")

def convert_ppt_files_in_dir(input_dir, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    for file in os.listdir(input_dir):
        if file.endswith(".ppt") or file.endswith(".pptx"):
            ppt_file = os.path.join(input_dir, file)
            file_name = os.path.splitext(file)[0]
            pdf_file = os.path.join(output_dir, file_name + ".pdf")
            convert(input_dir, output_dir)

# # Example usage
# input_dir = r"D:/tanma/Downloads/VIT Downloads/BCSE355L-AWS Solutions Architect/abc"  # Replace with the path where your PPTs are
# output_dir = r"D:/tanma/Downloads/VIT Downloads/BCSE355L-AWS Solutions Architect/abc/pdfs"  # Replace with the path where you want to save the PDFs

# convert_ppt_files_in_dir(input_dir, output_dir)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert PPT files to PDF.")
    parser.add_argument("input_dir", help="Directory containing PPT files to convert.")
    parser.add_argument("output_dir", help="Directory to save converted PDF files.")
    
    args = parser.parse_args()
    
    convert_ppt_files_in_dir(args.input_dir, args.output_dir)