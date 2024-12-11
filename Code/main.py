import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QFileDialog,
    QLabel, QTableWidget, QTableWidgetItem, QHBoxLayout, QMenu, QAction, QMessageBox
)
from PyQt5.QtCore import Qt, QPoint
from PyQt5.QtGui import QPalette, QColor
from fpdf import FPDF

class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Office 365 User Filter')
        self.setGeometry(100, 100, 1200, 800)

        # Initialize data
        self.data = None
        self.filtered_data = None
        self.column_menu = None  # Store menu as instance variable

        # Central widget
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        # Main layout
        self.main_layout = QVBoxLayout(self.central_widget)

        # Upload CSV button
        self.upload_button = QPushButton('Upload CSV')
        self.upload_button.clicked.connect(self.upload_file)
        self.main_layout.addWidget(self.upload_button)

        # Filter section
        self.filter_label = QLabel('Select Filters:')
        self.main_layout.addWidget(self.filter_label)

        # Horizontal layout for filter buttons
        self.filter_button_layout = QHBoxLayout()

        # Define filters
        self.filters = ['Active', 'Licensed', 'Unlicensed', 'Sign-In Blocked', 'Sign-In Allowed']
        self.filter_buttons = {}
        for filter_name in self.filters:
            button = QPushButton(filter_name)
            button.setCheckable(True)
            button.clicked.connect(self.apply_filters)
            self.filter_button_layout.addWidget(button)
            self.filter_buttons[filter_name] = button

        self.main_layout.addLayout(self.filter_button_layout)

        # Column selection section
        self.column_label = QLabel('Select Columns:')
        self.main_layout.addWidget(self.column_label)

        # Column selection button
        self.column_button = QPushButton('Select Columns')
        self.column_button.clicked.connect(self.show_column_menu)
        self.main_layout.addWidget(self.column_button)

        # List to keep track of selected columns
        self.selected_columns = []

        # Data table
        self.table = QTableWidget()
        self.main_layout.addWidget(self.table)

        # Export buttons layout
        self.export_button_layout = QHBoxLayout()

        # Export to CSV
        self.export_csv_button = QPushButton('Export to CSV')
        self.export_csv_button.clicked.connect(self.export_to_csv)
        self.export_button_layout.addWidget(self.export_csv_button)

        # Export to TXT
        self.export_txt_button = QPushButton('Export to TXT')
        self.export_txt_button.clicked.connect(self.export_to_txt)
        self.export_button_layout.addWidget(self.export_txt_button)

        # Export to PDF
        self.export_pdf_button = QPushButton('Export to PDF')
        self.export_pdf_button.clicked.connect(self.export_to_pdf)
        self.export_button_layout.addWidget(self.export_pdf_button)

        self.main_layout.addLayout(self.export_button_layout)

    def upload_file(self):
        """Handle CSV file upload."""
        file, _ = QFileDialog.getOpenFileName(self, 'Open CSV', '', 'CSV Files (*.csv)')
        if file:
            try:
                self.data = pd.read_csv(file)
                self.filtered_data = self.data.copy()

                # Reset filters
                for button in self.filter_buttons.values():
                    button.setChecked(False)

                # Select all columns by default
                self.selected_columns = self.data.columns.tolist()
                self.update_column_menu()

                # Display data
                self.display_data()

                QMessageBox.information(self, 'Success', 'CSV file loaded successfully.')
            except Exception as e:
                QMessageBox.critical(self, 'Error', f'Failed to load CSV file:\n{e}')

    def show_column_menu(self):
        """Show the column selection menu."""
        if self.data is None:
            QMessageBox.warning(self, 'Warning', 'Please upload a CSV file first.')
            return

        # Create a new dialog window
        dialog = QWidget(self)
        dialog.setWindowFlags(Qt.Window | Qt.WindowStaysOnTopHint)
        dialog.setWindowTitle('Select Columns')
        dialog.setGeometry(200, 200, 400, 500)
        
        # Main layout for dialog
        main_dialog_layout = QVBoxLayout()

        # Add Select All / Deselect All buttons
        select_buttons_layout = QHBoxLayout()
        select_all_btn = QPushButton('Select All')
        deselect_all_btn = QPushButton('Deselect All')
        select_buttons_layout.addWidget(select_all_btn)
        select_buttons_layout.addWidget(deselect_all_btn)
        main_dialog_layout.addLayout(select_buttons_layout)

        # Create scrollable area
        scroll = QWidget()
        scroll.setLayout(QVBoxLayout())
        
        # Create grid layout for checkboxes
        grid = QWidget()
        grid_layout = QHBoxLayout()
        left_column = QVBoxLayout()
        right_column = QVBoxLayout()
        
        # Create checkboxes for each column
        self.checkboxes = {}
        columns = list(self.data.columns)
        mid_point = len(columns) // 2
        
        for i, col in enumerate(columns):
            checkbox = QPushButton(col)
            checkbox.setCheckable(True)
            checkbox.setChecked(col in self.selected_columns)
            self.checkboxes[col] = checkbox
            
            if i < mid_point:
                left_column.addWidget(checkbox)
            else:
                right_column.addWidget(checkbox)

        grid_layout.addLayout(left_column)
        grid_layout.addLayout(right_column)
        grid.setLayout(grid_layout)
        scroll.layout().addWidget(grid)
        
        # Add scroll area to main layout
        main_dialog_layout.addWidget(scroll)

        # Add OK button
        ok_button = QPushButton('OK')
        ok_button.clicked.connect(dialog.close)
        main_dialog_layout.addWidget(ok_button)

        # Connect select/deselect all buttons
        def select_all():
            for checkbox in self.checkboxes.values():
                checkbox.setChecked(True)
            self.selected_columns = list(self.data.columns)
            self.display_data()

        def deselect_all():
            # Keep at least one column selected
            first_col = list(self.checkboxes.keys())[0]
            for col, checkbox in self.checkboxes.items():
                checkbox.setChecked(col == first_col)
            self.selected_columns = [first_col]
            self.display_data()

        select_all_btn.clicked.connect(select_all)
        deselect_all_btn.clicked.connect(deselect_all)

        # Connect individual checkboxes
        for col, checkbox in self.checkboxes.items():
            checkbox.clicked.connect(lambda checked, col=col: self.update_columns(col, checked))

        dialog.setLayout(main_dialog_layout)
        dialog.show()

    def update_columns(self, column, checked):
        """Update the selected columns based on user selection."""
        if checked and column not in self.selected_columns:
            self.selected_columns.append(column)
        elif not checked and column in self.selected_columns:
            if len(self.selected_columns) > 1:  # Ensure at least one column remains selected
                self.selected_columns.remove(column)
            else:
                QMessageBox.warning(self, 'Warning', 'At least one column must be selected.')
                # Re-check the button if we can't uncheck it
                self.checkboxes[column].setChecked(True)
                return

        self.display_data()

    def update_column_menu(self):
        """Ensure all columns are selected by default."""
        self.selected_columns = self.data.columns.tolist()

    def apply_filters(self):
        """Apply selected filters to the data."""
        if self.data is None:
            QMessageBox.warning(self, 'Warning', 'Please upload a CSV file first.')
            return

        self.filtered_data = self.data.copy()
        selected_filters = [name for name, btn in self.filter_buttons.items() if btn.isChecked()]

        # Define filter mappings: filter name to (column, condition function)
        filter_mappings = {
            'Active': ('State', lambda x: str(x).strip().lower() == 'active'),
            'Licensed': ('Licenses', lambda x: pd.notna(x) and str(x).strip() != ''),
            'Unlicensed': ('Licenses', lambda x: pd.isna(x) or str(x).strip() == ''),
            'Sign-In Blocked': ('Block credential', self.is_blocked),
            'Sign-In Allowed': ('Block credential', lambda x: not self.is_blocked(x))
        }

        for filter_name in selected_filters:
            if filter_name in filter_mappings:
                column, condition = filter_mappings[filter_name]
                actual_column = self.get_column_case_insensitive(column)
                if actual_column:
                    try:
                        self.filtered_data = self.filtered_data[self.filtered_data[actual_column].apply(condition)]
                    except Exception as e:
                        QMessageBox.warning(self, 'Warning', f"Error applying filter '{filter_name}': {e}")
                else:
                    QMessageBox.warning(
                        self,
                        'Warning',
                        f"Column '{column}' not found in data. Cannot apply filter '{filter_name}'."
                    )

        self.display_data()

    def is_blocked(self, value):
        """
        Determine if the user is sign-in blocked based on the 'Block credential' column.
        Adjust this function based on how the CSV represents blocked status.
        """
        if pd.isna(value):
            return False  # Assume not blocked if data is missing

        value_str = str(value).strip().lower()
        blocked_values = ['true', 'yes', '1']
        return value_str in blocked_values

    def get_column_case_insensitive(self, column_name):
        """Retrieve column name in a case-insensitive manner."""
        for col in self.data.columns:
            if col.strip().lower() == column_name.strip().lower():
                return col
        return None

    def display_data(self):
        """Display the filtered and selected data in the table."""
        if self.filtered_data is None:
            return

        try:
            if not self.selected_columns:
                QMessageBox.warning(self, 'Warning', 'No columns selected to display.')
                return

            display_data = self.filtered_data[self.selected_columns]

            self.table.setRowCount(display_data.shape[0])
            self.table.setColumnCount(display_data.shape[1])
            self.table.setHorizontalHeaderLabels(display_data.columns.tolist())
            self.table.setSortingEnabled(False)
            self.table.clearContents()

            for i in range(display_data.shape[0]):
                for j in range(display_data.shape[1]):
                    item = QTableWidgetItem(str(display_data.iat[i, j]))
                    self.table.setItem(i, j, item)

            self.table.setSortingEnabled(True)
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Error displaying data:\n{e}')

    def export_to_csv(self):
        """Export the displayed data to a CSV file."""
        if self.filtered_data is None or self.filtered_data.empty:
            QMessageBox.warning(self, 'Warning', 'No data to export.')
            return

        try:
            export_data = self.filtered_data[self.selected_columns]
            file, _ = QFileDialog.getSaveFileName(self, 'Save CSV', '', 'CSV Files (*.csv)')
            if file:
                export_data.to_csv(file, index=False)
                QMessageBox.information(self, 'Success', 'Data exported to CSV successfully.')
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Failed to export CSV:\n{e}')

    def export_to_txt(self):
        """Export the displayed data to a TXT file."""
        if self.filtered_data is None or self.filtered_data.empty:
            QMessageBox.warning(self, 'Warning', 'No data to export.')
            return

        try:
            export_data = self.filtered_data[self.selected_columns]
            file, _ = QFileDialog.getSaveFileName(self, 'Save TXT', '', 'Text Files (*.txt)')
            if file:
                with open(file, 'w') as f:
                    f.write(export_data.to_string(index=False))
                QMessageBox.information(self, 'Success', 'Data exported to TXT successfully.')
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Failed to export TXT:\n{e}')

    def export_to_pdf(self):
        """Export the displayed data to a PDF file."""
        if self.filtered_data is None or self.filtered_data.empty:
            QMessageBox.warning(self, 'Warning', 'No data to export.')
            return

        try:
            export_data = self.filtered_data[self.selected_columns]
            file, _ = QFileDialog.getSaveFileName(self, 'Save PDF', '', 'PDF Files (*.pdf)')
            if file:
                pdf = FPDF()
                pdf.add_page('L')  # Landscape orientation for more width
                
                # Title section with background
                pdf.set_fill_color(51, 122, 183)  # Professional blue
                pdf.rect(0, 0, pdf.w, 20, 'F')
                pdf.set_text_color(255, 255, 255)  # White text
                pdf.set_font("Arial", 'B', 16)
                pdf.cell(0, 15, "Office 365 User Report", 0, 1, 'C')
                
                # Reset text color and add some spacing
                pdf.set_text_color(0, 0, 0)
                pdf.ln(5)

                # Column headers with style
                pdf.set_font("Arial", 'B', 9)
                pdf.set_fill_color(240, 240, 240)  # Light gray background
                
                # Calculate column widths based on content
                col_widths = {}
                for col in export_data.columns:
                    # Get max width needed for column header
                    header_width = pdf.get_string_width(col) + 4  # Add padding
                    
                    # Get max width needed for column data
                    max_data_width = 0
                    for value in export_data[col]:
                        str_val = str(value) if pd.notna(value) else ''
                        max_data_width = max(max_data_width, pdf.get_string_width(str_val) + 4)
                    
                    # Use the larger of header or data width, with a minimum of 20 and maximum of 80
                    col_widths[col] = max(20, min(max(header_width, max_data_width), 120))
                
                # Print headers
                for col in export_data.columns:
                    pdf.cell(col_widths[col], 7, col, 1, 0, 'C', 1)
                pdf.ln()

                # Print rows with alternating background
                pdf.set_font("Arial", '', 8)
                for i, (_, row) in enumerate(export_data.iterrows()):
                    # Alternate row colors
                    if i % 2 == 0:
                        pdf.set_fill_color(255, 255, 255)  # White
                    else:
                        pdf.set_fill_color(245, 245, 245)  # Light gray
                    
                    # Store starting position
                    start_x = pdf.get_x()
                    start_y = pdf.get_y()
                    
                    # Calculate wrapped text and max height for entire row
                    row_cells = []
                    max_lines = 1
                    
                    # First pass: Calculate wrapped text for all cells
                    for col in export_data.columns:
                        value = row[col]
                        cell_text = str(value) if pd.notna(value) else ''
                        lines = []
                        current_line = ""
                        words = cell_text.split()
                        
                        for word in words:
                            test_line = current_line + " " + word if current_line else word
                            if pdf.get_string_width(test_line) < col_widths[col] - 2:
                                current_line = test_line
                            else:
                                if current_line:
                                    lines.append(current_line)
                                current_line = word
                        if current_line:
                            lines.append(current_line)
                            
                        if not lines:
                            lines = ['']
                            
                        row_cells.append(lines)
                        max_lines = max(max_lines, len(lines))
                    
                    row_height = max_lines * 5  # 5 points per line
                    
                    # Check if we need a new page
                    if pdf.get_y() + row_height > pdf.h - 20:
                        pdf.add_page('L')
                        # Repeat headers
                        pdf.set_font("Arial", 'B', 9)
                        pdf.set_fill_color(240, 240, 240)
                        for header_col in export_data.columns:
                            pdf.cell(col_widths[header_col], 7, header_col, 1, 0, 'C', 1)
                        pdf.ln()
                        pdf.set_font("Arial", '', 8)
                        start_y = pdf.get_y()
                    
                    # Second pass: Print all cells with uniform height
                    for col_idx, lines in enumerate(row_cells):
                        x = start_x + sum(col_widths[col] for col in 
                            list(export_data.columns)[:col_idx])
                        pdf.set_xy(x, start_y)
                        
                        # Pad shorter columns with empty lines to match max_lines
                        while len(lines) < max_lines:
                            lines.append('')
                            
                        # Join all lines with newline character
                        cell_text = '\n'.join(lines)
                        
                        # Print cell with consistent height
                        pdf.multi_cell(col_widths[export_data.columns[col_idx]], 
                            row_height/max_lines, cell_text, 1, 'L', 1)
                    
                    # Move to next row
                    pdf.set_y(start_y + row_height)

                pdf.output(file)
                QMessageBox.information(self, 'Success', 'Data exported to PDF successfully.')
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Failed to export PDF:\n{e}')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = App()
    window.show()
    sys.exit(app.exec_())
