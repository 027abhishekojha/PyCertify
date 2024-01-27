# PyCertify
 PyCertify automates certificate generation using PowerPoint templates &amp; Excel data. It replaces placeholders with participant names, creating both PPTX &amp; PDF certificates. Ideal for events, workshops, &amp; training sessions, PyCertify streamlines administrative tasks, ensuring efficient recognition of participant achievements.


 # Certificate Generator

This Python script allows you to generate certificates for participants using a PowerPoint template and data from an Excel sheet. It replaces placeholder text in the template with participant names and produces both PowerPoint and PDF versions of the certificates.

## Prerequisites

- Python 3.x
- Python libraries: `pandas`, `pptx`, `reportlab`

## Setup

1. Clone or download this repository to your local machine.
2. Install the required Python libraries using pip:
    ```bash
    pip install pandas python-pptx reportlab
    ```

## Usage

1. Ensure you have a PowerPoint template with a placeholder text (e.g., `<<NAME_PLACEHOLDER>>`) that will be replaced with participant names.
2. Prepare an Excel sheet (`participants.xlsx`) with a column named `ParticipantName` containing the names of the participants.
3. Modify the script by setting the following variables:
   - `template_path`: Path to your PowerPoint template file.
   - `excel_path`: Path to your Excel file containing participant names.
   - `output_folder`: Path to the folder where generated certificates will be saved.
   - `num_copies`: Number of certificates to generate.
4. Run the script by executing the following command:
    ```bash
    python certificate_generator.py
    ```

## Script Explanation

- `generate_certificates`: This function generates certificates by replacing placeholder text in the PowerPoint template with participant names and saving both PowerPoint and PDF versions of the certificates.
- `replace_text_in_shape`: This function replaces text in a PowerPoint shape.
- `calculate_run_position`: This function calculates the position of text runs in PowerPoint slides.
- The script uses `pandas` to read participant names from the Excel sheet and `pptx` to manipulate PowerPoint files. It also utilizes `reportlab` to create PDFs from PowerPoint slides.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---
Feel free to customize the instructions and details according to your specific project requirements. Let me know if you need further assistance!
