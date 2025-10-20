import xlwings as xw

def say_hello():
    """
    This function will be called from Excel.
    It prints 'Hello World' to a message box in Excel.
    """
    # Get the active workbook
    wb = xw.Book.caller()
    
    # Show a message box in Excel
    wb.app.alert("Hello World from Python! 🎉")
    
    # Also write to a cell (optional - just to show interaction)
    sheet = wb.sheets[0]
    sheet.range('A1').value = "Hello World from Python!"
    print("Hello World!")  # This prints to console/terminal if run from there

if __name__ == "__main__":
    # This allows the script to be run from Excel via xlwings
    xw.Book("hello_world.xlsm").set_mock_caller()
    say_hello()