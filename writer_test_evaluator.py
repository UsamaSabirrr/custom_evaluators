#TODO
"""
1- Bug, when some part is bold and other is not it still pass even though it should fail.
isse is if golden has bold in those sections and in attempt file i remove some bold even for those paragraphs it gives one. need to fix this.
"""

#!/usr/bin/env python3
"""
Test Suite for OSWorld DOCX Evaluator
Automatically tests all .docx files in test_data_evaluators folder
"""

import os
import sys
import subprocess
import urllib.request
from pathlib import Path
from datetime import datetime
import tempfile

# ============================================================================
# CONFIGURATION
# ============================================================================
TEST_DATA_DIR = Path("./docs_data")
GOLD_FILE_URL = "https://huggingface.co/datasets/Usamas3/osworld_tasks_files/resolve/main/MINI_PROJECT_STUDENT_TABLE_golden_v5.docx"
GOLD_FILE_PATH = "/tmp/golden_reference.docx"
TIMEOUT_SECONDS = 10

# Generate timestamp for log file
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
LOG_FILE = f"evaluator_test_results_{timestamp}.log"

# ============================================================================
# COLOR CODES FOR TERMINAL OUTPUT
# ============================================================================
class Colors:
    GREEN = '\033[92m'
    RED = '\033[91m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    BOLD = '\033[1m'
    RESET = '\033[0m'

# ============================================================================
# LOGGING HELPER
# ============================================================================
log_file_handle = None

def log_message(message, to_file=True, to_console=True):
    """Print message and optionally write to log file."""
    if to_console:
        print(message)
    if to_file and log_file_handle:
        log_file_handle.write(message + '\n')
        log_file_handle.flush()

# ============================================================================
# EMBEDDED EVALUATOR SCRIPT (CLEANED - PURE PYTHON)
# ============================================================================
# FIX NOTE: Removed 'bash -c' wrapper and cleaned up all escaped quotes.
EVALUATOR_SCRIPT = r"""
import sys
import os
import zipfile
import xml.etree.ElementTree as ET
import urllib.request

# ============================================================================
# CONFIGURATION - Set these variables for your specific use case
# ============================================================================
ACTUAL_FILE_PATH = '__ACTUAL_FILE_PATH__'  # Path to the solution file to evaluate
GOLDEN_FILE_URL = '__GOLDEN_FILE_URL__'  # URL to golden reference file
# ============================================================================

def normalize_font_name(font):
    if not font:
        return None
    font = font.replace(' (Body)', '').replace(' (Headings)', '')
    font = font.replace(' (body)', '').replace(' (headings)', '')
    font = font.strip()
    return font if font else None

def normalize_formatting(run_fmt):
    bold = run_fmt.get('bold')
    italic = run_fmt.get('italic')
    underline = run_fmt.get('underline')
    
    return {
        'bold': True if bold else False,
        'italic': True if italic else False,
        'underline': True if underline else False,
        'font_size': run_fmt.get('font_size'),
        'font_name': normalize_font_name(run_fmt.get('font_name'))
    }

def merge_consecutive_runs(runs_formatting):
    if not runs_formatting:
        return []
    
    merged = []
    current = runs_formatting[0].copy()
    
    for run in runs_formatting[1:]:
        current_fmt = normalize_formatting(current)
        run_fmt = normalize_formatting(run)
        
        if (current_fmt['bold'] == run_fmt['bold'] and
            current_fmt['italic'] == run_fmt['italic'] and
            current_fmt['underline'] == run_fmt['underline'] and
            current_fmt['font_size'] == run_fmt['font_size'] and
            current_fmt['font_name'] == run_fmt['font_name']):
            current['text'] = current.get('text', '') + run.get('text', '')
        else:
            if current.get('text'):
                merged.append(current)
            current = run.copy()
    
    if current.get('text'):
        merged.append(current)
    
    return merged

def extract_paragraphs_with_full_details(docx_path):
    paragraphs = []
    try:
        with zipfile.ZipFile(docx_path, 'r') as docx:
            ns = {
                'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
            }
            doc_xml = docx.read('word/document.xml')
            root = ET.fromstring(doc_xml)
            
            for para in root.findall('.//w:p', ns):
                style_elem = para.find('.//w:pStyle', ns)
                style = style_elem.get('{' + ns['w'] + '}val') if style_elem is not None else None
                
                pPr = para.find('.//w:pPr', ns)
                alignment = None
                spacing_before = None
                spacing_after = None
                line_spacing = None
                
                if pPr is not None:
                    jc = pPr.find('.//w:jc', ns)
                    if jc is not None:
                        alignment = jc.get('{' + ns['w'] + '}val')
                    
                    spacing = pPr.find('.//w:spacing', ns)
                    if spacing is not None:
                        spacing_before = spacing.get('{' + ns['w'] + '}before')
                        spacing_after = spacing.get('{' + ns['w'] + '}after')
                        line_spacing = spacing.get('{' + ns['w'] + '}line')
                
                text_parts = []
                runs_formatting = []
                
                for run in para.findall('.//w:r', ns):
                    run_text = ''
                    for text_elem in run.findall('.//w:t', ns):
                        if text_elem.text:
                            run_text += text_elem.text
                    
                    rPr = run.find('.//w:rPr', ns)
                    bold = False
                    italic = False
                    underline = False
                    font_size = None
                    font_name = None
                    
                    if rPr is not None:
                        bold = rPr.find('.//w:b', ns) is not None
                        italic = rPr.find('.//w:i', ns) is not None
                        underline = rPr.find('.//w:u', ns) is not None
                        
                        sz = rPr.find('.//w:sz', ns)
                        if sz is not None:
                            font_size = sz.get('{' + ns['w'] + '}val')
                        
                        rFonts = rPr.find('.//w:rFonts', ns)
                        if rFonts is not None:
                            font_name = rFonts.get('{' + ns['w'] + '}ascii')
                    
                    if run_text:
                        text_parts.append(run_text)
                        runs_formatting.append({
                            'text': run_text,
                            'bold': bold,
                            'italic': italic,
                            'underline': underline,
                            'font_size': font_size,
                            'font_name': font_name
                        })
                
                text = ''.join(text_parts)
                
                has_comment = (
                    para.find('.//w:commentRangeStart', ns) is not None or
                    para.find('.//w:commentReference', ns) is not None
                )
                
                paragraphs.append({
                    'text': text,
                    'style': style,
                    'alignment': alignment,
                    'spacing_before': spacing_before,
                    'spacing_after': spacing_after,
                    'line_spacing': line_spacing,
                    'runs_formatting': runs_formatting,
                    'has_comment': has_comment
                })
                
    except Exception as e:
        print(f"ERROR extracting paragraphs: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc(file=sys.stderr)
        return None
    
    return paragraphs

def extract_comments(docx_path):
    comments = []
    try:
        with zipfile.ZipFile(docx_path, 'r') as docx:
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            
            try:
                comments_xml = docx.read('word/comments.xml')
                root = ET.fromstring(comments_xml)
                
                for comment in root.findall('.//w:comment', ns):
                    author = comment.get('{' + ns['w'] + '}author', '')
                    comment_id = comment.get('{' + ns['w'] + '}id', '')
                    
                    text_parts = []
                    for text_elem in comment.findall('.//w:t', ns):
                        if text_elem.text:
                            text_parts.append(text_elem.text)
                    
                    text = ''.join(text_parts).strip()
                    
                    comments.append({
                        'id': comment_id,
                        'author': author,
                        'text': text
                    })
            except KeyError:
                pass
                
    except Exception as e:
        print(f"ERROR extracting comments: {e}", file=sys.stderr)
        return None
    
    return comments

def extract_headers_footers(docx_path):
    headers_footers = {'headers': [], 'footers': []}
    
    try:
        with zipfile.ZipFile(docx_path, 'r') as docx:
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            file_list = docx.namelist()
            
            for filename in file_list:
                if filename.startswith('word/header'):
                    header_xml = docx.read(filename)
                    root = ET.fromstring(header_xml)
                    text_parts = []
                    for text_elem in root.findall('.//w:t', ns):
                        if text_elem.text:
                            text_parts.append(text_elem.text)
                    if text_parts:
                        headers_footers['headers'].append(''.join(text_parts))
                
                elif filename.startswith('word/footer'):
                    footer_xml = docx.read(filename)
                    root = ET.fromstring(footer_xml)
                    text_parts = []
                    for text_elem in root.findall('.//w:t', ns):
                        if text_elem.text:
                            text_parts.append(text_elem.text)
                    if text_parts:
                        headers_footers['footers'].append(''.join(text_parts))
                        
    except Exception as e:
        print(f"ERROR extracting headers/footers: {e}", file=sys.stderr)
        return None
    
    return headers_footers

def extract_document_structure(docx_path):
    structure = {
        'tables': 0,
        'images': 0,
        'sections': 0,
        'footnotes': 0,
        'endnotes': 0,
        'bookmarks': 0
    }
    
    try:
        with zipfile.ZipFile(docx_path, 'r') as docx:
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                 'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture'}
            doc_xml = docx.read('word/document.xml')
            root = ET.fromstring(doc_xml)
            
            structure['tables'] = len(root.findall('.//w:tbl', ns))
            structure['images'] = len(root.findall('.//pic:pic', ns))
            structure['sections'] = len(root.findall('.//w:sectPr', ns))
            
            bookmarks = root.findall('.//w:bookmarkStart', ns)
            structure['bookmarks'] = len([b for b in bookmarks 
                                           if not b.get('{' + ns['w'] + '}name', '').startswith('_')])
            
            file_list = docx.namelist()
            if 'word/footnotes.xml' in file_list:
                footnotes_xml = docx.read('word/footnotes.xml')
                footnotes_root = ET.fromstring(footnotes_xml)
                footnotes = footnotes_root.findall('.//w:footnote', ns)
                structure['footnotes'] = len([f for f in footnotes 
                                               if f.get('{' + ns['w'] + '}type') != 'separator'])
            
            if 'word/endnotes.xml' in file_list:
                endnotes_xml = docx.read('word/endnotes.xml')
                endnotes_root = ET.fromstring(endnotes_xml)
                endnotes = endnotes_root.findall('.//w:endnote', ns)
                structure['endnotes'] = len([e for e in endnotes 
                                              if e.get('{' + ns['w'] + '}type') != 'separator'])
                
    except Exception as e:
        print(f"ERROR extracting structure: {e}", file=sys.stderr)
        return None
    
    return structure

def compare_runs_semantic(actual_runs, golden_runs, para_index):
    actual_merged = merge_consecutive_runs(actual_runs)
    golden_merged = merge_consecutive_runs(golden_runs)
    
    actual_merged = [r for r in actual_merged if r.get('text', '').strip()]
    golden_merged = [r for r in golden_merged if r.get('text', '').strip()]
    
    if len(actual_merged) != len(golden_merged):
        actual_full_text = ''.join(r['text'] for r in actual_merged)
        golden_full_text = ''.join(r['text'] for r in golden_merged)
        
        if actual_full_text.strip() == golden_full_text.strip():
            if len(actual_merged) > 0 and len(golden_merged) > 0:
                actual_uniform = all(
                    normalize_formatting(r) == normalize_formatting(actual_merged[0])
                    for r in actual_merged
                )
                golden_uniform = all(
                    normalize_formatting(r) == normalize_formatting(golden_merged[0])
                    for r in golden_merged
                )
                
                if actual_uniform and golden_uniform:
                    actual_fmt = normalize_formatting(actual_merged[0])
                    golden_fmt = normalize_formatting(golden_merged[0])
                    if actual_fmt == golden_fmt:
                        return True, None, None
        
        return (False, 'FAIL_PARAGRAPH_RUN_COUNT_CHANGED',
                f'Para {para_index}: Expected {len(golden_merged)} formatted segments, got {len(actual_merged)} (after merging consecutive runs with same formatting)')
    
    for j, (actual_run, golden_run) in enumerate(zip(actual_merged, golden_merged)):
        actual_norm = normalize_formatting(actual_run)
        golden_norm = normalize_formatting(golden_run)
        
        actual_text_clean = actual_run['text'].strip()
        golden_text_clean = golden_run['text'].strip()
        
        if actual_text_clean != golden_text_clean:
            return (False, 'FAIL_RUN_TEXT_CHANGED',
                    f'Para {para_index}, Run {j}: Expected text "{golden_text_clean}", got "{actual_text_clean}"')
        
        if actual_norm['bold'] != golden_norm['bold']:
            return (False, 'FAIL_RUN_BOLD_CHANGED',
                    f'Para {para_index}, Run {j}: Bold mismatch (expected {golden_norm["bold"]}, got {actual_norm["bold"]})')
        
        if actual_norm['italic'] != golden_norm['italic']:
            return (False, 'FAIL_RUN_ITALIC_CHANGED',
                    f'Para {para_index}, Run {j}: Italic mismatch (expected {golden_norm["italic"]}, got {actual_norm["italic"]})')
        
        if actual_norm['underline'] != golden_norm['underline']:
            return (False, 'FAIL_RUN_UNDERLINE_CHANGED',
                    f'Para {para_index}, Run {j}: Underline mismatch (expected {golden_norm["underline"]}, got {actual_norm["underline"]})')
        
        if actual_norm['font_size'] != golden_norm['font_size']:
            return (False, 'FAIL_RUN_FONT_SIZE_CHANGED',
                    f'Para {para_index}, Run {j}: Font size mismatch (expected {golden_norm["font_size"]}, got {actual_norm["font_size"]})')
        
        if actual_norm['font_name'] != golden_norm['font_name']:
            return (False, 'FAIL_RUN_FONT_NAME_CHANGED',
                    f'Para {para_index}, Run {j}: Font name mismatch (expected {golden_norm["font_name"]}, got {actual_norm["font_name"]})')
    
    return True, None, None

try:
    # Use the configured paths
    doc_path = ACTUAL_FILE_PATH
    golden_url = GOLDEN_FILE_URL
    golden_path = '/tmp/golden_reference.docx'
    
    if not os.path.exists(doc_path):
        print('FAIL_FILE_NOT_FOUND', end='')
        print(f'Solution file not found at: {doc_path}', file=sys.stderr)
        sys.exit(0)
    
    print(f'DEBUG: Evaluating file: {doc_path}', file=sys.stderr)
    print(f'DEBUG: Downloading golden reference from: {golden_url}', file=sys.stderr)
    
    urllib.request.urlretrieve(golden_url, golden_path)
    print('DEBUG: Golden reference downloaded', file=sys.stderr)
    
    print('\n=== EXTRACTING DOCUMENTS ===', file=sys.stderr)
    
    actual_paragraphs = extract_paragraphs_with_full_details(doc_path)
    golden_paragraphs = extract_paragraphs_with_full_details(golden_path)
    
    if actual_paragraphs is None or golden_paragraphs is None:
        print('FAIL_PARAGRAPH_EXTRACTION', end='')
        sys.exit(0)
    
    print('\n=== CHECK 1: PARAGRAPH COUNT ===', file=sys.stderr)
    
    if len(actual_paragraphs) != len(golden_paragraphs):
        print('FAIL_PARAGRAPH_COUNT_CHANGED', end='')
        print(f'Expected {len(golden_paragraphs)} paragraphs, got {len(actual_paragraphs)}', file=sys.stderr)
        sys.exit(0)
    
    print(f'✓ Paragraph count matches: {len(actual_paragraphs)}', file=sys.stderr)
    
    print('\n=== CHECK 2: PARAGRAPH TEXT AND FORMATTING ===', file=sys.stderr)
    
    for i in range(len(actual_paragraphs)):
        actual_para = actual_paragraphs[i]
        golden_para = golden_paragraphs[i]
        
        actual_text_normalized = ' '.join(actual_para['text'].split())
        golden_text_normalized = ' '.join(golden_para['text'].split())
        
        if actual_text_normalized != golden_text_normalized:
            print('FAIL_PARAGRAPH_TEXT_CHANGED', end='')
            print(f'Para {i}: Expected text "{golden_text_normalized}", got "{actual_text_normalized}"', file=sys.stderr)
            sys.exit(0)
        
        if actual_para['style'] != golden_para['style']:
            print('FAIL_PARAGRAPH_STYLE_CHANGED', end='')
            print(f'Para {i}: Expected style "{golden_para["style"]}", got "{actual_para["style"]}"', file=sys.stderr)
            sys.exit(0)
        
        if actual_para['alignment'] != golden_para['alignment']:
            print('FAIL_PARAGRAPH_ALIGNMENT_CHANGED', end='')
            print(f'Para {i}: Expected alignment "{golden_para["alignment"]}", got "{actual_para["alignment"]}"', file=sys.stderr)
            sys.exit(0)
        
        if actual_para['spacing_before'] != golden_para['spacing_before']:
            print('FAIL_PARAGRAPH_SPACING_BEFORE_CHANGED', end='')
            print(f'Para {i}: Expected spacing_before "{golden_para["spacing_before"]}", got "{actual_para["spacing_before"]}"', file=sys.stderr)
            sys.exit(0)
        
        if actual_para['spacing_after'] != golden_para['spacing_after']:
            print('FAIL_PARAGRAPH_SPACING_AFTER_CHANGED', end='')
            print(f'Para {i}: Expected spacing_after "{golden_para["spacing_after"]}", got "{actual_para["spacing_after"]}"', file=sys.stderr)
            sys.exit(0)
        
        if actual_para['line_spacing'] != golden_para['line_spacing']:
            print('FAIL_PARAGRAPH_LINE_SPACING_CHANGED', end='')
            print(f'Para {i}: Expected line_spacing "{golden_para["line_spacing"]}", got "{actual_para["line_spacing"]}"', file=sys.stderr)
            sys.exit(0)
        
        is_match, error_code, error_msg = compare_runs_semantic(
            actual_para['runs_formatting'],
            golden_para['runs_formatting'],
            i
        )
        
        if not is_match:
            print(error_code, end='')
            if error_msg:
                print(error_msg, file=sys.stderr)
            sys.exit(0)
    
    print(f'✓ All {len(actual_paragraphs)} paragraphs match (text and formatting)', file=sys.stderr)
    
    print('\n=== CHECK 3: HEADERS AND FOOTERS ===', file=sys.stderr)
    
    actual_hf = extract_headers_footers(doc_path)
    golden_hf = extract_headers_footers(golden_path)
    
    if actual_hf is None or golden_hf is None:
        print('FAIL_HEADER_FOOTER_EXTRACTION', end='')
        sys.exit(0)
    
    if actual_hf['headers'] != golden_hf['headers']:
        print('FAIL_HEADERS_CHANGED', end='')
        print(f'Expected headers: {golden_hf["headers"]}, got: {actual_hf["headers"]}', file=sys.stderr)
        sys.exit(0)
    
    if actual_hf['footers'] != golden_hf['footers']:
        print('FAIL_FOOTERS_CHANGED', end='')
        print(f'Expected footers: {golden_hf["footers"]}, got: {actual_hf["footers"]}', file=sys.stderr)
        sys.exit(0)
    
    print('✓ Headers and footers match', file=sys.stderr)
    
    print('\n=== CHECK 4: DOCUMENT STRUCTURE ===', file=sys.stderr)
    
    actual_structure = extract_document_structure(doc_path)
    golden_structure = extract_document_structure(golden_path)
    
    if actual_structure is None or golden_structure is None:
        print('FAIL_STRUCTURE_EXTRACTION', end='')
        sys.exit(0)
    
    if actual_structure['tables'] != golden_structure['tables']:
        print('FAIL_TABLE_COUNT_CHANGED', end='')
        print(f'Expected {golden_structure["tables"]} tables, got {actual_structure["tables"]}', file=sys.stderr)
        sys.exit(0)
    
    if actual_structure['images'] != golden_structure['images']:
        print('FAIL_IMAGE_COUNT_CHANGED', end='')
        print(f'Expected {golden_structure["images"]} images, got {actual_structure["images"]}', file=sys.stderr)
        sys.exit(0)
    
    if actual_structure['sections'] != golden_structure['sections']:
        print('FAIL_SECTION_COUNT_CHANGED', end='')
        print(f'Expected {golden_structure["sections"]} sections, got {actual_structure["sections"]}', file=sys.stderr)
        sys.exit(0)
    
    if actual_structure['bookmarks'] != golden_structure['bookmarks']:
        print('FAIL_BOOKMARK_COUNT_CHANGED', end='')
        print(f'Expected {golden_structure["bookmarks"]} bookmarks, got {actual_structure["bookmarks"]}', file=sys.stderr)
        sys.exit(0)
    
    if actual_structure['footnotes'] != golden_structure['footnotes']:
        print('FAIL_FOOTNOTE_COUNT_CHANGED', end='')
        print(f'Expected {golden_structure["footnotes"]} footnotes, got {actual_structure["footnotes"]}', file=sys.stderr)
        sys.exit(0)
    
    if actual_structure['endnotes'] != golden_structure['endnotes']:
        print('FAIL_ENDNOTE_COUNT_CHANGED', end='')
        print(f'Expected {golden_structure["endnotes"]} endnotes, got {actual_structure["endnotes"]}', file=sys.stderr)
        sys.exit(0)
    
    print('✓ Document structure matches', file=sys.stderr)
    
    print('\n=== CHECK 5: COMMENTS ===', file=sys.stderr)
    
    actual_comments = extract_comments(doc_path)
    golden_comments = extract_comments(golden_path)
    
    if actual_comments is None or golden_comments is None:
        print('FAIL_COMMENT_EXTRACTION', end='')
        sys.exit(0)
    
    if len(actual_comments) != len(golden_comments):
        print('FAIL_COMMENT_COUNT_CHANGED', end='')
        print(f'Expected {len(golden_comments)} comments, got {len(actual_comments)}', file=sys.stderr)
        sys.exit(0)
    
    for i, (actual_comment, golden_comment) in enumerate(zip(actual_comments, golden_comments)):
        if actual_comment['text'] != golden_comment['text']:
            print('FAIL_COMMENT_TEXT_CHANGED', end='')
            print(f'Comment {i}: Expected "{golden_comment["text"]}", got "{actual_comment["text"]}"', file=sys.stderr)
            sys.exit(0)
        
        if actual_comment['author'] != golden_comment['author']:
            print('FAIL_COMMENT_AUTHOR_CHANGED', end='')
            print(f'Comment {i}: Expected author "{golden_comment["author"]}", got "{actual_comment["author"]}"', file=sys.stderr)
            sys.exit(0)
    
    # Check comment placement
    for i, (actual_para, golden_para) in enumerate(zip(actual_paragraphs, golden_paragraphs)):
        if actual_para['has_comment'] != golden_para['has_comment']:
            print('FAIL_COMMENT_PLACEMENT_CHANGED', end='')
            print(f'Para {i}: Comment placement mismatch (expected has_comment={golden_para["has_comment"]}, got {actual_para["has_comment"]})', file=sys.stderr)
            sys.exit(0)
    
    print(f'✓ All {len(actual_comments)} comments match (text, author, and placement)', file=sys.stderr)
    
    print('\n=== ALL CHECKS PASSED ===', file=sys.stderr)
    print('SUCCESS', end='')
    
except Exception as e:
    print('FAIL_EXCEPTION', end='')
    print(f'Exception: {e}', file=sys.stderr)
    import traceback
    traceback.print_exc(file=sys.stderr)
"""
# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def download_golden_file():
    """Download golden reference file from HuggingFace."""
    try:
        log_message(f"\n{Colors.BLUE}Downloading golden file from HuggingFace...{Colors.RESET}")
        log_message(f"URL: {GOLD_FILE_URL}")
        
        urllib.request.urlretrieve(GOLD_FILE_URL, GOLD_FILE_PATH)
        
        if os.path.exists(GOLD_FILE_PATH):
            file_size = os.path.getsize(GOLD_FILE_PATH)
            log_message(f"{Colors.GREEN}✓ Golden file downloaded successfully{Colors.RESET}")
            log_message(f"  Location: {GOLD_FILE_PATH}")
            log_message(f"  Size: {file_size:,} bytes")
            return True
        else:
            log_message(f"{Colors.RED}✗ Download failed - file not found{Colors.RESET}")
            return False
            
    except Exception as e:
        log_message(f"{Colors.RED}✗ Error downloading golden file: {e}{Colors.RESET}")
        return False

def find_test_files():
    """Auto-discover all .docx files in test data directory."""
    if not TEST_DATA_DIR.exists():
        log_message(f"{Colors.RED}✗ Test data directory not found: {TEST_DATA_DIR}{Colors.RESET}")
        return []
    
    test_files = sorted(TEST_DATA_DIR.glob("*.docx"))
    
    if not test_files:
        log_message(f"{Colors.YELLOW}⚠ No .docx files found in {TEST_DATA_DIR}{Colors.RESET}")
        return []
    
    log_message(f"\n{Colors.BOLD}Found {len(test_files)} test file(s):{Colors.RESET}")
    for f in test_files:
        log_message(f"  - {f.name}")
    
    return test_files

def run_evaluator(test_file_path):
    """Run evaluator on a test file and return results."""
    # FIX NOTE: Using .replace() instead of .format() to avoid KeyError
    evaluator_code = EVALUATOR_SCRIPT.replace('__ACTUAL_FILE_PATH__', str(test_file_path.absolute()))
    evaluator_code = evaluator_code.replace('__GOLDEN_FILE_URL__', GOLD_FILE_URL)
    
    # Write to temp file
    with tempfile.NamedTemporaryFile(mode='w', suffix='.py', delete=False) as f:
        f.write(evaluator_code)
        temp_script = f.name
    
    try:
        # Run evaluator
        result = subprocess.run(
            ['python3', temp_script],
            capture_output=True,
            text=True,
            timeout=TIMEOUT_SECONDS
        )
        
        return {
            'stdout': result.stdout,
            'stderr': result.stderr,
            'returncode': result.returncode,
            'timeout': False
        }
        
    except subprocess.TimeoutExpired:
        return {
            'stdout': '',
            'stderr': f'Evaluator timed out after {TIMEOUT_SECONDS} seconds',
            'returncode': -1,
            'timeout': True
        }
    finally:
        # Clean up temp file
        try:
            os.unlink(temp_script)
        except:
            pass

# ============================================================================
# MAIN TEST SUITE
# ============================================================================

def main():
    global log_file_handle
    
    # Open log file
    log_file_handle = open(LOG_FILE, 'w')
    
    # Print header
    header = f"""
{'=' * 80}
{Colors.BOLD}OSWorld DOCX Evaluator Test Suite{Colors.RESET}
{'=' * 80}
"""
    log_message(header)
    
    # Test metadata
    start_time = datetime.now()
    log_message(f"Test started at: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    log_message(f"Test data directory: {TEST_DATA_DIR.absolute()}")
    log_message(f"Golden file URL: {GOLD_FILE_URL}")
    log_message(f"Log file: {LOG_FILE}")
    
    # Download golden file
    log_message(f"\n{'=' * 80}")
    if not download_golden_file():
        log_message(f"\n{Colors.RED}FATAL ERROR: Could not download golden file{Colors.RESET}")
        log_message(f"{'=' * 80}\n")
        log_file_handle.close()
        return 1
    
    # Find test files
    log_message(f"\n{'=' * 80}")
    test_files = find_test_files()
    
    if not test_files:
        log_message(f"\n{Colors.YELLOW}No test files to process{Colors.RESET}")
        log_message(f"{'=' * 80}\n")
        log_file_handle.close()
        return 1
    
    # Run tests
    log_message(f"\n{'=' * 80}\n")
    
    results = []
    passed = 0
    failed = 0
    
    for idx, test_file in enumerate(test_files, 1):
        log_message(f"{Colors.BOLD}[Test {idx}/{len(test_files)}] {test_file.name}{Colors.RESET}")
        log_message(f"File path: {test_file.absolute()}")
        
        # Run evaluator
        result = run_evaluator(test_file)
        
        # Determine verdict
        is_pass = result['stdout'].strip() == 'SUCCESS' and result['returncode'] == 0
        
        if is_pass:
            passed += 1
            log_message(f"Result: {Colors.GREEN}✓ PASS{Colors.RESET} - Output is VALID")
        else:
            failed += 1
            error_code = result['stdout'].strip() if result['stdout'].strip() else 'UNKNOWN_ERROR'
            log_message(f"Result: {Colors.RED}✗ FAIL{Colors.RESET}")
            log_message(f"  Error Code: {error_code}")
            
            # Show detailed error logs from stderr
            if result['stderr']:
                log_message(f"\n{Colors.YELLOW}  Detailed Error Log:{Colors.RESET}")
                for line in result['stderr'].split('\n'):
                    if line.strip():
                        log_message(f"    {line}")
            
            if result['timeout']:
                log_message(f"  {Colors.RED}Evaluator timed out!{Colors.RESET}")
        
        log_message('-' * 80 + '\n')
        
        results.append({
            'file': test_file.name,
            'passed': is_pass,
            'error_code': result['stdout'].strip() if not is_pass else None,
            'stderr': result['stderr']
        })
    
    # Summary
    log_message(f"{'=' * 80}")
    log_message(f"{Colors.BOLD}TEST SUMMARY{Colors.RESET}")
    log_message(f"{'=' * 80}\n")
    
    total = len(test_files)
    pass_rate = (passed / total * 100) if total > 0 else 0
    
    log_message(f"Total Tests: {total}")
    log_message(f"Passed: {Colors.GREEN}{passed}{Colors.RESET}")
    log_message(f"Failed: {Colors.RED}{failed}{Colors.RESET}")
    log_message(f"Pass Rate: {pass_rate:.1f}%")
    
    log_message(f"\n{'=' * 80}")
    if failed > 0:
        log_message(f"{Colors.RED}{Colors.BOLD}⚠ {failed} TEST(S) FAILED - CHECK RESULTS ABOVE{Colors.RESET}")
    else:
        log_message(f"{Colors.GREEN}{Colors.BOLD}✓ ALL TESTS PASSED{Colors.RESET}")
    log_message(f"{'=' * 80}\n")
    
    end_time = datetime.now()
    duration = (end_time - start_time).total_seconds()
    log_message(f"Test completed at: {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
    log_message(f"Total duration: {duration:.2f} seconds")
    log_message(f"\nComplete log saved to: {Colors.BOLD}{LOG_FILE}{Colors.RESET}\n")
    
    log_file_handle.close()
    
    return 0 if failed == 0 else 1

if __name__ == "__main__":
    sys.exit(main())