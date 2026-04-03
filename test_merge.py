from playwright.sync_api import sync_playwright
import time

def test_merge_functionality():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        
        print("Navigating to http://localhost:8080...")
        page.goto('http://localhost:8080')
        page.wait_for_load_state('networkidle')
        
        print("Taking initial screenshot...")
        page.screenshot(path='d:/AI-project/eXcel-html/test_screenshots/01_initial.png', full_page=True)
        
        print("\n=== Testing Tab Switching ===")
        convert_tab = page.locator('[data-tab="convert"]')
        merge_tab = page.locator('[data-tab="merge"]')
        
        print("Checking if convert tab is active by default...")
        assert convert_tab.is_visible(), "Convert tab button should be visible"
        assert merge_tab.is_visible(), "Merge tab button should be visible"
        
        print("Clicking merge tab...")
        merge_tab.click()
        page.wait_for_timeout(500)
        
        print("Taking screenshot after clicking merge tab...")
        page.screenshot(path='d:/AI-project/eXcel-html/test_screenshots/02_merge_tab.png', full_page=True)
        
        print("\n=== Testing Merge Tab Elements ===")
        
        merge_drop_zone = page.locator('#mergeDropZone')
        assert merge_drop_zone.is_visible(), "Merge drop zone should be visible"
        
        merge_files_input = page.locator('#mergeFiles')
        assert merge_files_input.is_visible(), "Merge files input should be visible"
        
        validate_btn = page.locator('#validateBtn')
        assert validate_btn.is_visible(), "Validate button should be visible"
        
        merge_btn = page.locator('#mergeBtn')
        assert merge_btn.is_visible(), "Merge button should be visible"
        assert merge_btn.is_disabled(), "Merge button should be disabled initially"
        
        print("\n=== Testing Configuration Options ===")
        
        enable_dedup_label = page.locator('label:has(#enableDedup)')
        enable_dedup = page.locator('#enableDedup')
        assert enable_dedup.count() > 0, "Enable dedup checkbox should exist"
        assert enable_dedup.is_checked(), "Enable dedup should be checked by default"
        
        dedup_config = page.locator('#dedupConfig')
        assert dedup_config.is_visible(), "Dedup config should be visible when dedup is enabled"
        
        print("Unchecking dedup option...")
        enable_dedup_label.click()
        page.wait_for_timeout(300)
        assert not enable_dedup.is_checked(), "Enable dedup should be unchecked"
        
        print("Re-checking dedup option...")
        enable_dedup_label.click()
        page.wait_for_timeout(300)
        assert enable_dedup.is_checked(), "Enable dedup should be checked again"
        
        preserve_format = page.locator('#preserveFormat')
        assert preserve_format.count() > 0, "Preserve format checkbox should exist"
        assert preserve_format.is_checked(), "Preserve format should be checked by default"
        
        skip_empty_rows = page.locator('#skipEmptyRows')
        assert skip_empty_rows.count() > 0, "Skip empty rows checkbox should exist"
        assert skip_empty_rows.is_checked(), "Skip empty rows should be checked by default"
        
        print("\n=== Testing Clear All Files Button ===")
        
        clear_all_btn = page.locator('#clearAllFiles')
        assert clear_all_btn.count() > 0, "Clear all files button should exist"
        assert not clear_all_btn.is_visible(), "Clear all files button should be hidden when no files"
        
        print("\n=== Testing Modal Functionality ===")
        
        validate_btn.click()
        page.wait_for_timeout(500)
        
        error_modal = page.locator('#errorModal')
        assert error_modal.is_visible(), "Error modal should be visible when validating with no files"
        
        page.screenshot(path='d:/AI-project/eXcel-html/test_screenshots/03_error_modal.png', full_page=True)
        
        modal_ok_btn = page.locator('#modalOkBtn')
        modal_ok_btn.click()
        page.wait_for_timeout(300)
        assert not error_modal.is_visible(), "Error modal should be hidden after clicking OK"
        
        print("\n=== Testing Tab Switch Back ===")
        convert_tab.click()
        page.wait_for_timeout(500)
        
        convert_tab_content = page.locator('#convert-tab')
        assert convert_tab_content.is_visible(), "Convert tab content should be visible"
        
        page.screenshot(path='d:/AI-project/eXcel-html/test_screenshots/04_convert_tab.png', full_page=True)
        
        print("\n=== All Tests Passed! ===")
        
        browser.close()

if __name__ == "__main__":
    import os
    os.makedirs('d:/AI-project/eXcel-html/test_screenshots', exist_ok=True)
    test_merge_functionality()
