import re
from odf.opendocument import OpenDocumentText
from odf.text import P

def analyze_ini_file(ini_file_path):
    data = {}
    current_section = None
    
    # Read and analyze the .ini file
    with open(ini_file_path, 'r') as file:
        for line in file:
            # Remove whitespace and newlines
            line = line.strip()
            
            # Match sections like [abc] or [abc:children]
            section_match = re.match(r'\[(.+)\]', line)
            if section_match:
                current_section = section_match.group(1)
                if current_section not in data:
                    data[current_section] = 0
            else:
                # Count each line under the current section
                if current_section:
                    data[current_section] += 1

    return data

def write_to_odt(results, output_file_path):
    # Create ODT document
    textdoc = OpenDocumentText()

    # Add each result as a paragraph
    for section, count in results.items():
        p = P(text=f"Section [{section}] has {count} related lines.")
        textdoc.text.addElement(p)

    # Save the document
    textdoc.save(output_file_path)

def main():
    ini_file_path = 'example.ini'  # Your .ini file path
    output_file_path = 'analysis_results.odt'  # Output ODT file path

    # Analyze the .ini file
    results = analyze_ini_file(ini_file_path)

    # Write the analysis results to an ODT file
    write_to_odt(results, output_file_path)
    print(f"Results saved to {output_file_path}")

if __name__ == '__main__':
    main()
