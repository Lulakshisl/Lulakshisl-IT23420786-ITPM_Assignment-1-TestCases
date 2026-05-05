from playwright.sync_api import sync_playwright
import time
import os
import argparse
import re
from pathlib import Path
import unicodedata
import openpyxl

# ---------- Configuration ----------
DEFAULT_FRONTEND_URL = "https://www.pixelssuite.com/chat-translator"

# ---------- Helper Functions ----------
def normalize_text(text):
    if not isinstance(text, str):
        return ""
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def get_cell_value(ws, row, col):
    try:
        cell = ws.cell(row=row, column=col)
        return cell.value
    except:
        return None

def set_cell_value(ws, row, col, value):
    try:
        ws.cell(row=row, column=col).value = value
    except:
        pass

def get_translation(page, input_text, wait_time=25):
    """Get fresh translation for each input"""
    try:
        # Find elements
        textareas = page.locator("textarea").all()
        if len(textareas) >= 2:
            input_area = textareas[0]
            output_area = textareas[1]
        else:
            input_area = page.locator("textarea").first
            output_area = page.locator("textarea").last
        
        # Step 1: Clear output area first (by clicking Clear button if available)
        try:
            clear_btn = page.get_by_role("button", name=re.compile("clear", re.IGNORECASE))
            if clear_btn.is_visible(timeout=2000):
                clear_btn.click()
                time.sleep(1)
        except:
            pass
        
        # Step 2: Clear input area completely
        input_area.click()
        time.sleep(0.3)
        
        # Select all and delete
        page.keyboard.press("Control+A")
        time.sleep(0.2)
        page.keyboard.press("Delete")
        time.sleep(0.5)
        
        # Fill with empty string to ensure clearing
        input_area.fill("")
        time.sleep(0.5)
        
        # Step 3: Type new input
        input_area.fill(input_text)
        time.sleep(1)
        
        # Verify text was entered
        entered_text = input_area.input_value()
        if entered_text != input_text:
            # Retry typing
            input_area.fill("")
            time.sleep(0.5)
            input_area.fill(input_text)
            time.sleep(1)
        
        # Step 4: Click Translate button
        translate_btn = page.get_by_role("button", name=re.compile("transliterate", re.IGNORECASE))
        translate_btn.click()
        
        # Step 5: Wait for translation to complete
        time.sleep(wait_time)
        
        # Step 6: Get output
        output = output_area.input_value()
        if not output:
            output = output_area.inner_text()
        
        return output.strip() if output else ""
        
    except Exception as e:
        print(f"      Error: {e}")
        return ""

def refresh_page(page, url):
    """Refresh the page to get a clean state"""
    page.reload()
    time.sleep(3)
    
    # Handle any popups
    try:
        accept_btn = page.get_by_role("button", name=re.compile("accept", re.IGNORECASE))
        if accept_btn.is_visible(timeout=3000):
            accept_btn.click()
            time.sleep(1)
    except:
        pass

def run_test():
    parser = argparse.ArgumentParser()
    parser.add_argument("--excel", required=True)
    parser.add_argument("--url", default=DEFAULT_FRONTEND_URL)
    parser.add_argument("--wait", type=int, default=20, help="Wait time in seconds")
    parser.add_argument("--refresh-every", type=int, default=5, help="Refresh page every N tests")
    parser.add_argument("--headless", action="store_true")
    args = parser.parse_args()

    # Load Excel
    if not os.path.exists(args.excel):
        print(f"❌ Excel not found: {args.excel}")
        return

    print(f"\n{'='*70}")
    print(f"📊 Loading: {os.path.basename(args.excel)}")
    
    wb = openpyxl.load_workbook(args.excel)
    ws = wb["Test cases"]
    
    # Find column indices
    header_row = 1
    input_col = None
    expected_col = None
    actual_col = None
    status_col = None
    type_col = None
    tc_id_col = None
    
    for col in range(1, 20):
        val = ws.cell(row=header_row, column=col).value
        if val:
            val_str = str(val).strip().lower()
            if val_str == 'input':
                input_col = col
            elif val_str == 'expected output':
                expected_col = col
            elif val_str == 'actual output':
                actual_col = col
            elif val_str == 'status':
                status_col = col
            elif val_str == 'input length type':
                type_col = col
            elif val_str == 'test case id':
                tc_id_col = col
    
    print(f"✅ Input column: {input_col}")
    print(f"✅ Expected column: {expected_col}")
    print(f"✅ Actual column: {actual_col}")
    print(f"✅ Status column: {status_col}")
    
    # Find all test rows
    test_rows = []
    for row in range(2, ws.max_row + 1):
        input_val = get_cell_value(ws, row, input_col)
        if input_val and str(input_val).strip() and str(input_val).strip() != 'None':
            test_rows.append(row)
    
    print(f"📊 Found {len(test_rows)} test cases")
    print(f"{'='*70}\n")
    
    # Launch browser
    with sync_playwright() as p:
        print(f"🌐 Launching Chrome...")
        browser = p.chromium.launch(headless=args.headless, slow_mo=100)
        page = browser.new_page()
        
        print(f"🌐 Loading website...")
        page.goto(args.url)
        time.sleep(3)
        
        # Handle initial popups
        try:
            accept_btn = page.get_by_role("button", name=re.compile("accept", re.IGNORECASE))
            if accept_btn.is_visible(timeout=3000):
                accept_btn.click()
                time.sleep(1)
        except:
            pass
        
        print(f"✅ Ready! Starting {len(test_rows)} tests...\n")
        print(f"{'─'*70}")
        
        # Process each test
        for idx, row in enumerate(test_rows, 1):
            # Refresh page periodically to clear cache
            if idx > 1 and (idx - 1) % args.refresh_every == 0:
                print(f"\n🔄 Refreshing page (clean slate for test {idx})...")
                refresh_page(page, args.url)
                time.sleep(2)
            
            # Get test data
            tc_id = get_cell_value(ws, row, tc_id_col) if tc_id_col else f"Row{row}"
            length_type = get_cell_value(ws, row, type_col) if type_col else "?"
            input_text = str(get_cell_value(ws, row, input_col)).strip()
            expected_text = str(get_cell_value(ws, row, expected_col)).strip() if expected_col else ""
            
            print(f"\n[{idx}/{len(test_rows)}] {tc_id} | Row {row} | Type: {length_type}")
            print(f"📝 Input: {input_text[:70]}...")
            print(f"🔄 Translating... (waiting up to {args.wait}s)")
            
            # Get translation for this specific input
            actual_output = get_translation(page, input_text, wait_time=args.wait)
            
            # Determine status
            if not actual_output:
                status = "NO_OUTPUT"
                print(f"⚠️  No output received!")
            elif not expected_text:
                status = "NO_EXPECTED"
            else:
                # Compare normalized text
                if normalize_text(actual_output) == normalize_text(expected_text):
                    status = "PASS"
                else:
                    status = "FAIL"
            
            # Save to Excel
            if actual_col:
                set_cell_value(ws, row, actual_col, actual_output if actual_output else "No output received")
            if status_col:
                set_cell_value(ws, row, status_col, status)
            
            print(f"📤 Output: {actual_output[:80] if actual_output else '[EMPTY]'}")
            print(f"📊 Status: {status}")
            
            # Save every 3 tests
            if idx % 3 == 0 or idx == len(test_rows):
                wb.save(args.excel)
                print(f"💾 Saved ({idx}/{len(test_rows)})")
            
            # Wait between tests
            time.sleep(2)
        
        # Final save
        wb.save(args.excel)
        
        # Print summary
        print(f"\n{'='*70}")
        print(f"📊 TEST SUMMARY")
        print(f"{'='*70}")
        
        # Count results
        pass_count = 0
        fail_count = 0
        no_output = 0
        
        for row in test_rows:
            status = get_cell_value(ws, row, status_col)
            if status == "PASS":
                pass_count += 1
            elif status == "FAIL":
                fail_count += 1
            elif status == "NO_OUTPUT":
                no_output += 1
        
        print(f"✅ PASS: {pass_count}")
        print(f"❌ FAIL: {fail_count}")
        print(f"⚠️  NO_OUTPUT: {no_output}")
        print(f"{'='*70}")
        print(f"📁 Results saved to: {args.excel}")
        print(f"{'='*70}\n")
        
        # Keep browser open if needed
        if not args.headless:
            input("Press Enter to close browser...")
        
        browser.close()

if __name__ == "__main__":
    run_test()