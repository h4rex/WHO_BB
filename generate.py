import pandas as pd
from datetime import datetime
import os
import fitz  # PyMuPDF
from barcode import Code39
from barcode.writer import ImageWriter
import time
import logging
import json
import concurrent.futures
import argparse
from functools import partial

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Constants
BARCODE_WAIT_TIMEOUT = 5  # seconds
BARCODE_WAIT_INTERVAL = 0.1  # seconds
BARCODE_RECT = fitz.Rect(81, 18, 270, 50)
TEXT_INSERTIONS = {
    'BESTELLING_POSITIE': (70, 84, 7.5),
    'ONTVANGSDATUM': (70, 94, 7.5),
}

class LabelCreationError(Exception):
    pass

def load_config(config_file):
    """Load configuration from a file and validate."""
    try:
        with open(config_file, 'r') as f:
            config = json.load(f)
        # Validate config
        required_keys = ['trucklist_file', 'template_folder', 'output_folder', 'barcode_folder', 'font']
        for key in required_keys:
            if key not in config:
                raise ValueError(f"Missing required config key: {key}")
        return config
    except Exception as e:
        logging.error(f"Error loading configuration: {e}")
        raise

def setup_directories(config):
    """Ensure necessary directories exist."""
    try:
        os.makedirs(config['barcode_folder'], exist_ok=True)
        os.makedirs(config['output_folder'], exist_ok=True)
    except Exception as e:
        logging.error(f"Error setting up directories: {e}")
        raise

def extract_trucklist_info(trucklist_file):
    """Extract necessary information from the trucklist."""
    try:
        df = pd.read_excel(trucklist_file)
        orders_to_process = df[df['Labels created?'] == 'No']
        return orders_to_process
    except Exception as e:
        logging.error(f"Error extracting trucklist info: {e}")
        raise

def generate_barcode(order_number, position, barcode_folder, font_path):
    """Generate a barcode using the python-barcode library."""
    try:
        barcode_data = f"00{order_number}{position}"
        barcode = Code39(barcode_data, writer=ImageWriter(), add_checksum=False)
        barcode_filename = f'barcode_{order_number}_{position}'
        barcode_path = os.path.join(barcode_folder, barcode_filename)
        
        font_path = os.path.join(os.getcwd(), font_path)
        barcode.save(barcode_path, {
            'module_width': 0.2, 
            'module_height': 6.0, 
            'quiet_zone': 0, 
            'font_size': 7.5, 
            'text_distance': 3.0,
            'font_path': font_path  # Set the font path
        })
        return f"{barcode_path}.png"
    except Exception as e:
        logging.error(f"Error generating barcode: {e}")
        raise LabelCreationError(f"Failed to generate barcode: {e}")

def create_label(order_info, template_folder, output_folder, barcode_folder, font_path):
    """Create a label using a PDF template."""
    try:
        sku = order_info['SKU']
        sku_full = sku
        template_subfolder = os.path.join(template_folder, sku[:5])

        if not os.path.isdir(template_subfolder):
            raise FileNotFoundError(f"Template subfolder not found: {template_subfolder}")

        template_file = None
        for file in os.listdir(template_subfolder):
            if file.startswith(f"Labels {sku_full}") and file.endswith(".pdf"):
                template_file = os.path.join(template_subfolder, file)
                break

        if template_file is None:
            raise FileNotFoundError(f"Template file not found for SKU: {sku_full} in {template_subfolder}")

        logging.info(f"Using template: {template_file}")

        order_id = str(order_info['Ext order number'])
        position = f"00{int(float(order_info['position'])):03d}00"
        order_date = order_info['Order date']
        if isinstance(order_date, pd.Timestamp):
            order_date = order_date.strftime('%d.%m.%Y')
        else:
            order_date = datetime.strptime(order_date, '%d.%m.%Y').strftime('%d.%m.%Y')

        pdf_document = fitz.open(template_file)

        barcode_image_path = generate_barcode(order_id, position, barcode_folder, font_path)
        start_time = time.time()
        while not os.path.exists(barcode_image_path):
            if time.time() - start_time > BARCODE_WAIT_TIMEOUT:
                raise FileNotFoundError(f"Barcode image file not found: {barcode_image_path}")
            time.sleep(BARCODE_WAIT_INTERVAL)

        page = pdf_document[0]
        page.insert_image(BARCODE_RECT, filename=barcode_image_path)

        bestelling_positie = f"{order_id}/{int(position):05d}"
        text_insertions = {
            'BESTELLING_POSITIE': (70, 84, bestelling_positie, 7.5),
            'ONTVANGSDATUM': (70, 94, order_date, 7.5),
        }

        for key, (x, y, text, font_size) in text_insertions.items():
            page.insert_text((x, y), text, fontsize=font_size, fontname="helv", fill=(0, 0, 0))

        order_folder = os.path.join(output_folder, order_id)
        os.makedirs(order_folder, exist_ok=True)
        
        label_filename = os.path.abspath(os.path.join(order_folder, f"{sku}_{position}.pdf"))
        pdf_document.save(label_filename)
        pdf_document.close()

        os.remove(barcode_image_path)

        return label_filename
    except Exception as e:
        logging.error(f"Error creating label: {e}")
        raise LabelCreationError(f"Failed to create label: {e}")

def update_trucklist(trucklist_file, orders_processed):
    """Update the trucklist after processing."""
    try:
        df = pd.read_excel(trucklist_file)
        df.loc[df['Order number'].isin(orders_processed), 'Labels created?'] = 'Yes'
        df.to_excel(trucklist_file, index=False)
    except Exception as e:
        logging.error(f"Error updating trucklist: {e}")
        raise

def process_order(order_info, template_folder, output_folder, barcode_folder, font_path):
    """Process a single order."""
    try:
        label_filename = create_label(order_info, template_folder, output_folder, barcode_folder, font_path)
        if label_filename:
            logging.info(f"Label created: {label_filename}")
            return order_info['Order number']
        else:
            logging.error(f"Failed to create label for order {order_info['Order number']}")
    except LabelCreationError as e:
        logging.error(e)
    except Exception as e:
        logging.error(f"Failed to create label for order {order_info['Order number']}: {e}")
    return None

def main(config_file):
    try:
        config = load_config(config_file)
        setup_directories(config)
        trucklist_file = os.path.abspath(config['trucklist_file'])
        
        logging.info(f"Current directory: {os.getcwd()}")
        logging.info(f"Files in trucklist directory: {os.listdir(os.path.dirname(trucklist_file))}")
        logging.info(f"Files in template directory: {os.listdir(config['template_folder'])}")
        
        orders_to_process = extract_trucklist_info(trucklist_file)
        
        process_order_partial = partial(process_order, template_folder=config['template_folder'], output_folder=config['output_folder'], barcode_folder=config['barcode_folder'], font_path=config['font'])

        with concurrent.futures.ThreadPoolExecutor() as executor:
            results = list(executor.map(process_order_partial, [order for _, order in orders_to_process.iterrows()]))
        
        orders_processed = [order for order in results if order]

        update_trucklist(trucklist_file, orders_processed)
    except Exception as e:
        logging.error(f"Error in main execution: {e}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Process label creation.')
    parser.add_argument('--config', type=str, default='config.json', help='Path to the configuration file.')
    args = parser.parse_args()
    main(args.config)
    input("Press Enter to continue...")  # This line will pause the script
