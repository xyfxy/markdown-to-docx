import sys
sys.path.append('/app') # Ensure md_parser can be imported

from md_parser import parse_markdown
import os

if __name__ == '__main__':
    test_file = "test_document.md"
    
    # Check if md_parser module and parse_markdown function are available
    try:
        from md_parser import parse_markdown
        print("Successfully imported parse_markdown from md_parser.")
    except ImportError as e:
        print(f"Error importing md_parser: {e}")
        exit(1)
    except AttributeError as e:
        print(f"Error: parse_markdown function not found in md_parser: {e}")
        exit(1)

    print(f"--- Parsing {test_file} ---")
    parsed_elements = parse_markdown(test_file)
    if parsed_elements and parsed_elements[0].get('type') == 'error':
        print(f"Error during parsing: {parsed_elements[0]['text']}")
    else:
        for element in parsed_elements:
            print(element)
    print("--- End of Parsing ---")

    # Clean up the dummy markdown file
    try:
        os.remove(test_file)
        print(f"Cleaned up {test_file}")
    except OSError as e:
        print(f"Error cleaning up {test_file}: {e}")
