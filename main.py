import click
import os
import logging # Added for logging
import traceback # Added for stack trace formatting
from md_parser import parse_markdown # Assuming this can raise exceptions
from docx_generator import DocxWriter # Assuming this can raise exceptions
from typing import List, Dict, Any # Keep for existing type hints

# --- Logger Setup ---
# Configure logger. The level will be set in the CLI command function.
logger = logging.getLogger(__name__)
# Basic configuration will be done in the CLI command.
# A common practice is to configure it once at the application entry point.

# --- List Styling Definitions ---
UNORDERED_LIST_BULLETS = ["• ", "◦ ", "▪ "]
ORDERED_LIST_PREFIX_FORMAT = "{}. "

@click.command(help="Converts a Markdown file to a DOCX document with specified formatting.")
@click.argument('input_md', type=click.Path(exists=True, dir_okay=False, readable=True))
@click.argument('output_docx', type=click.Path(writable=True, dir_okay=False)) # Click handles basic writability check
@click.option('--line-spacing', 'line_spacing_pt', type=float, default=None, help='Body text line spacing in points (e.g., 27.0).')
@click.option('--font-size', 'font_size_pt', type=float, default=None, help='Body text font size in points (e.g., 16.0 for 3号).')
@click.option('--template', 'template_path', type=click.Path(exists=True, dir_okay=False, readable=True), default=None, help='Path to an optional Word template file (.docx).')
@click.option('--verbose', '-v', is_flag=True, default=False, help="Enable verbose output and detailed error stack traces.")
def markdown_to_docx_cli(input_md: str, output_docx: str, 
                         line_spacing_pt: float, font_size_pt: float, 
                         template_path: str, verbose: bool):
    """
    Converts the INPUT_MD Markdown file to OUTPUT_DOCX Word document.
    Formatting options for line spacing, font size, and a document template can be specified.
    """
    # --- Logging Configuration ---
    log_level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=log_level,
        format="%(asctime)s - %(levelname)s - %(name)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )
    # If you want to ensure only your app's logger is affected, configure specifically:
    # logger.setLevel(log_level)
    # if not logger.handlers: # Avoid adding multiple handlers on re-runs in some environments
    #     ch = logging.StreamHandler()
    #     ch.setLevel(log_level)
    #     formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(name)s - %(message)s", "%Y-%m-%d %H:%M:%S")
    #     ch.setFormatter(formatter)
    #     logger.addHandler(ch)
    # logger.propagate = False # Prevent passing to root logger if configured separately

    logger.info(f"Starting conversion of '{input_md}' to '{output_docx}'...")
    if template_path:
        logger.info(f"Using template: '{template_path}'")
    if line_spacing_pt is not None:
        logger.info(f"Custom line spacing: {line_spacing_pt} pt")
    if font_size_pt is not None:
        logger.info(f"Custom body font size: {font_size_pt} pt")
    if verbose:
        logger.debug("Verbose mode enabled.")

    try:
        logger.debug("Instantiating DocxWriter...")
        writer = DocxWriter(
            output_docx_path=output_docx,
            line_spacing_pt=line_spacing_pt,
            font_size_pt=font_size_pt,
            template_path=template_path
            # Consider passing logger to DocxWriter if it needs to log internally:
            # logger=logger 
        )

        logger.debug(f"Parsing Markdown file: '{input_md}'...")
        parsed_elements = parse_markdown(input_md) # This function already handles FileNotFoundError for input_md
        
        if not parsed_elements:
            logger.warning("Markdown parsing resulted in no elements.")
            writer.save() # Save an empty document as per previous logic
            logger.info(f"Empty document saved to '{output_docx}'.")
            click.secho("Warning: Markdown parsing resulted in no elements. Empty document saved.", fg="yellow")
            return
        
        # md_parser returns a list with an error dict if it fails internally (e.g. file not found)
        if isinstance(parsed_elements, list) and parsed_elements and parsed_elements[0].get('type') == 'error':
            error_msg = parsed_elements[0]['text']
            logger.error(f"Error during Markdown parsing: {error_msg}")
            click.secho(f"Markdown Parsing Error: {error_msg}", fg="red")
            # This typically means input_md wasn't found by parse_markdown, already specific.
            return

        logger.info("Processing parsed Markdown elements...")
        list_level_counters: Dict[int, int] = {} 
        current_list_type_stack: List[str] = [] 

        # Heuristic for recipient
        if parsed_elements and parsed_elements[0]['type'] == 'paragraph':
            first_p_text = parsed_elements[0]['text']
            if first_p_text.endswith("：") or first_p_text.endswith(":"):
                recipient_text = first_p_text[:-1] 
                logger.debug(f"Attempting to add recipient: '{recipient_text}'")
                writer.add_recipient(recipient_text)
                logger.info(f"Processed first paragraph as recipient: '{recipient_text}'")
                parsed_elements = parsed_elements[1:]

        for i, element in enumerate(parsed_elements):
            el_type = element.get('type')
            el_text = element.get('text', '')
            logger.debug(f"Processing element {i+1}/{len(parsed_elements)}: type='{el_type}', text='{el_text[:30]}...'")

            if el_type in ['h1', 'h2', 'h3', 'h4', 'h5']:
                level_num = int(el_type[1:])
                writer.add_heading(el_text, level=level_num)
            elif el_type == 'paragraph':
                writer.add_paragraph(el_text)
            elif el_type == 'emphasis':
                writer.add_emphasis(el_text, style=element.get('style', 'italic'))
            elif el_type == 'ul_start':
                current_list_type_stack.append('ul')
                list_level_counters[len(current_list_type_stack) - 1] = 0 
            elif el_type == 'ol_start':
                current_list_type_stack.append('ol')
                list_level_counters[len(current_list_type_stack) - 1] = 0
            elif el_type == 'list_item':
                if not current_list_type_stack:
                    logger.warning(f"Encountered list_item '{el_text}' without active list context. Skipping.")
                    click.secho(f"Warning: List item '{el_text[:30]}...' found out of list context. Skipping.", fg="yellow")
                    continue
                current_level = element.get('level', len(current_list_type_stack) - 1)
                is_ordered = current_list_type_stack[current_level] == 'ol'
                prefix = ""
                if is_ordered:
                    list_level_counters[current_level] = list_level_counters.get(current_level, 0) + 1
                    prefix = ORDERED_LIST_PREFIX_FORMAT.format(list_level_counters[current_level])
                else:
                    prefix = UNORDERED_LIST_BULLETS[current_level % len(UNORDERED_LIST_BULLETS)]
                writer.add_list_item(el_text, ordered=is_ordered, level=current_level, item_prefix=prefix)
            elif el_type == 'ul_end' or el_type == 'ol_end':
                if current_list_type_stack:
                    current_list_type_stack.pop()
                else:
                    logger.warning(f"Encountered {el_type} without active list context. Ignoring.")
                    click.secho(f"Warning: {el_type} found out of list context. Ignoring.", fg="yellow")

        logger.info("Saving the document...")
        writer.save()
        logger.info(f"Successfully converted '{input_md}' to '{output_docx}'!")
        click.secho(f"Successfully converted '{input_md}' to '{output_docx}'!", fg="green")

    except FileNotFoundError as e:
        # This specifically catches if template_path is not found by DocxWriter,
        # or if input_md path check by Click fails (though Click handles its own message).
        # md_parser handles its own input_md FileNotFoundError.
        logger.error(f"File not found: {e.filename}. Details: {e.strerror}")
        click.secho(f"Error: File not found - {e.filename}. Please check the path and permissions.", fg="red", err=True)
        if verbose:
            click.echo(traceback.format_exc(), err=True)
    except PermissionError as e:
        logger.error(f"Permission denied for file: {e.filename}. Details: {e.strerror}")
        click.secho(f"Error: Permission denied - {e.filename}. Cannot read/write file. Please check permissions.", fg="red", err=True)
        if verbose:
            click.echo(traceback.format_exc(), err=True)
    # Placeholder for specific markdown parsing exceptions if known
    # except markdown.MarkdownException as e: # Replace with actual exception type
    #     logger.error(f"Markdown processing error: {str(e)}")
    #     click.secho(f"Error: Could not process Markdown content. Details: {str(e)}", fg="red", err=True)
    #     if verbose:
    #         click.echo(traceback.format_exc(), err=True)
    # Placeholder for specific python-docx exceptions if known
    # except docx.oxml.exceptions.OxmlError as e: # Replace with actual or broader docx exception
    #     logger.error(f"DOCX generation error: {str(e)}")
    #     click.secho(f"Error: Failed to generate DOCX file. Details: {str(e)}", fg="red", err=True)
    #     if verbose:
    #         click.echo(traceback.format_exc(), err=True)
    except Exception as e:
        # General catch-all for other unexpected errors
        error_name = e.__class__.__name__
        logger.error(f"An unexpected error occurred: {error_name} - {str(e)}")
        click.secho(f"An unexpected error ({error_name}) occurred: {str(e)}", fg="red", err=True)
        # Always print traceback for unexpected errors if verbose, or consider always for these.
        if verbose:
            click.echo(traceback.format_exc(), err=True)
        else:
            click.secho("Run with --verbose for more details.", fg="yellow", err=True)

if __name__ == '__main__':
    markdown_to_docx_cli()
```
