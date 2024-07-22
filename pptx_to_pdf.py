import os
import comtypes.client

# Function to convert pptx to pdf using comtypes
def ppt_to_pdf(input_file):
    pdf_file = input_file.replace('.pptx', '.pdf')
    try:
        print(f"Attempting to convert {input_file} to {pdf_file}")
        
        if not os.path.exists(input_file):
            print(f"File does not exist: {input_file}")
            return
        
        powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
        powerpoint.Visible = 1
        deck = powerpoint.Presentations.Open(input_file, WithWindow=False)
        deck.SaveAs(pdf_file, 32)  # 32 is the formatType for ppt to pdf
        deck.Close()
        powerpoint.Quit()
        
        # Verify the PDF was created before deleting the original pptx file
        if os.path.exists(pdf_file):
            os.remove(input_file)
            print(f"Converted {input_file} to {pdf_file} and deleted the original .pptx file")
        else:
            print(f"Failed to create PDF for {input_file}")
    except Exception as e:
        print(f"Error converting {input_file} to PDF: {e}")

# Function to get all pptx files in the folder and its subfolders
def get_pptx_files(folder):
    pptx_files = []
    for subdir, _, files in os.walk(folder):
        for file in files:
            if file.endswith('.pptx'):
                pptx_files.append(os.path.join(subdir, file))
    return pptx_files

# Main function to process templates sequentially
def main():
    template_folder = os.path.abspath('labels')
    
    print(f"Current directory: {os.getcwd()}")
    print(f"Files in template directory: {os.listdir(template_folder)}")

    try:
        pptx_files = get_pptx_files(template_folder)
        
        if not pptx_files:
            print("No .pptx files found in the specified directory.")
            return
        
        # Process files sequentially
        for pptx_file in pptx_files:
            ppt_to_pdf(pptx_file)
        
        print("Conversion completed.")
    except Exception as e:
        print(f"Failed to process templates: {e}")

if __name__ == "__main__":
    main()
