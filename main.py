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

        # Process elements with special handling for title and recipient
        title_found = False
        recipient_processed = False
        current_paragraph_obj = None # For merging paragraph and inline_code
        
        for i, element in enumerate(parsed_elements):
            el_type = element.get('type')
            el_text = element.get('text', '') # Default for elements that have text
            logger.debug(f"Processing element {i+1}/{len(parsed_elements)}: type='{el_type}', content='{str(element)[:100]}...'")

            if el_type in ['h1', 'h2', 'h3', 'h4', 'h5']:
                level_num = int(el_type[1:])
                writer.add_heading(el_text, level=level_num)
                if level_num == 1:
                    title_found = True
                    logger.debug(f"Processed main title: '{el_text}'")
                current_paragraph_obj = None # Reset for block elements
            elif el_type == 'paragraph':
                # Check if this should be treated as recipient
                if title_found and not recipient_processed:
                    logger.debug(f"Attempting to add recipient: '{el_text}'")
                    current_paragraph_obj = writer.add_recipient(el_text) # add_recipient also returns a paragraph obj
                    logger.info(f"Processed paragraph after title as recipient: '{el_text}'")
                    recipient_processed = True
                    # Recipient is a special block, subsequent inline code shouldn't merge with it.
                    # However, if add_recipient itself creates a full paragraph, then current_paragraph_obj is fine.
                    # For safety, and assuming recipient is a standalone block:
                    # current_paragraph_obj = None # Uncomment if recipient should not be merged with subsequent inline_code
                else:
                    logger.debug(f"Adding paragraph: '{el_text[:50]}...'")
                    current_paragraph_obj = writer.add_paragraph(el_text)
            elif el_type == 'inline_code':
                logger.debug(f"Adding inline_code: '{el_text}'")
                if current_paragraph_obj is None:
                    # Create a new paragraph if inline_code is not preceded by a paragraph element
                    # This handles cases where inline_code might appear after a heading, list, or at the start.
                    # According to PRD, inline code should be part of a paragraph.
                    # If md_parser ensures inline_code is always within a conceptual paragraph (even if not explicit <p>),
                    # this new paragraph creation might be for standalone inline code snippets.
                    # For now, we'll create a new paragraph if there isn't an active one.
                    logger.debug("No active paragraph for inline_code, creating a new one.")
                    current_paragraph_obj = writer.document.add_paragraph()
                     # Remove default first line indent for paragraphs created just for inline code
                    current_paragraph_obj.paragraph_format.first_line_indent = None
                writer.add_inline_code(el_text, current_paragraph_obj)
            elif el_type == 'emphasis':
                # Emphasis creates its own paragraph as per current docx_generator.py
                logger.debug(f"Adding emphasis: '{el_text}', style: {element.get('style')}")
                writer.add_emphasis(el_text, style=element.get('style', 'italic'))
                current_paragraph_obj = None # Reset for block elements
            elif el_type == 'ul_start' or el_type == 'ol_start':
                logger.debug(f"Starting list: {el_type}")
                current_list_type_stack.append('ul' if el_type == 'ul_start' else 'ol')
                list_level_counters[len(current_list_type_stack) - 1] = 0 
                current_paragraph_obj = None # Reset for block elements
            elif el_type == 'list_item':
                if not current_list_type_stack:
                    logger.warning(f"Encountered list_item '{el_text}' without active list context. Skipping.")
                    click.secho(f"Warning: List item '{el_text[:30]}...' found out of list context. Skipping.", fg="yellow")
                    current_paragraph_obj = None # Reset for safety
                    continue
                current_level = element.get('level', len(current_list_type_stack) - 1)
                is_ordered = current_list_type_stack[current_level] == 'ol'
                prefix = ""
                if is_ordered:
                    list_level_counters[current_level] = list_level_counters.get(current_level, 0) + 1
                    prefix = ORDERED_LIST_PREFIX_FORMAT.format(list_level_counters[current_level])
                else:
                    prefix = UNORDERED_LIST_BULLETS[current_level % len(UNORDERED_LIST_BULLETS)]
                logger.debug(f"Adding list_item: '{el_text[:30]}...', prefix: '{prefix}', level: {current_level}")
                writer.add_list_item(el_text, ordered=is_ordered, level=current_level, item_prefix=prefix)
                current_paragraph_obj = None # List items are block-like, reset current paragraph
            elif el_type == 'ul_end' or el_type == 'ol_end':
                logger.debug(f"Ending list: {el_type}")
                if current_list_type_stack:
                    current_list_type_stack.pop()
                else:
                    logger.warning(f"Encountered {el_type} without active list context. Ignoring.")
                    click.secho(f"Warning: {el_type} found out of list context. Ignoring.", fg="yellow")
                current_paragraph_obj = None # Reset for block elements
            # --- New Element Handling ---
            elif el_type == 'table':
                table_data = element # The whole element is the table data structure from md_parser
                logger.info(f"Adding table: {len(table_data.get('headers',[]))} header(s), {len(table_data.get('rows',[]))} row(s)")
                writer.add_table(table_data)
                current_paragraph_obj = None 
            elif el_type == 'code_block':
                logger.info(f"Adding code_block (lang: {element.get('language')}): {el_text[:50]}...")
                writer.add_code_block(el_text, element.get('language'))
                current_paragraph_obj = None
            elif el_type == 'blockquote':
                logger.info(f"Adding blockquote: {el_text[:50]}...")
                writer.add_blockquote(el_text)
                current_paragraph_obj = None
            elif el_type == 'link':
                logger.info(f"Adding link: Text='{el_text}', URL='{element.get('href')}'")
                writer.add_hyperlink_paragraph(el_text, element.get('href'))
                current_paragraph_obj = None
            elif el_type == 'image':
                src = element.get('src')
                alt = element.get('alt')
                logger.info(f"Adding image: src='{src}', alt='{alt}'")
                # Using a default width for images as discussed, e.g., 15cm.
                # This can be made configurable if needed.
                writer.add_image(src, alt, width_cm=15.0) 
                current_paragraph_obj = None
            elif el_type == 'horizontal_rule':
                logger.info("Adding horizontal_rule.")
                writer.add_horizontal_rule()
                current_paragraph_obj = None
            # --- End of New Element Handling ---
            else:
                logger.warning(f"Unknown element type encountered: '{el_type}'. Content: '{str(element)[:100]}'. Skipping.")
                click.secho(f"Warning: Unknown element type '{el_type}'. Skipping.", fg="yellow")
                current_paragraph_obj = None # Reset for unknown elements too

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
