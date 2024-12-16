import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QComboBox, QPushButton,
    QTabWidget, QMessageBox, QTableView
)
from PyQt5.QtCore import Qt, QAbstractTableModel
import pandas as pd
import math


# Utility functions
def calculate_with_tolerance(value, tolerance):
    """Calculate minimum and maximum value based on tolerance."""
    min_value = value - (value * tolerance / 100)
    max_value = value + (value * tolerance / 100)
    return min_value, max_value

def get_units(component_type):
    """Return a list of units based on the component type."""
    units_map = {
        "Capacitor": ["F", "mF", "µF", "uF", "nF", "pF"],
        "Resistor": ["Ω", "ohm", "mΩ", "kΩ", "MΩ"],
        "Inductor": ["H", "mH", "µH", "uH", "nH", "kH"]
    }
    return units_map.get(component_type, [])

def convert_units(value, from_unit, to_unit):
    """Convert between different units."""
    unit_factors = {
        "F": 1, "mF": 1e-3, "µF": 1e-6, "uF": 1e-6, "nF": 1e-9, "pF": 1e-12,
        "Ω": 1, "ohm": 1, "mΩ": 1e-3, "kΩ": 1e3, "MΩ": 1e6,
        "H": 1, "mH": 1e-3, "µH": 1e-6, "uH": 1e-6, "nH": 1e-9, "kH": 1e3
    }
    return value * unit_factors[from_unit] / unit_factors[to_unit]

# Pandas Model for QTableView
class PandasModel(QAbstractTableModel):
    def __init__(self, data):
        super().__init__()
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            return str(self._data.iloc[index.row(), index.column()])

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                return self._data.columns[section]
            if orientation == Qt.Vertical:
                return str(self._data.index[section])

# Component Calculator Tab
class ComponentCalculator(QWidget):
    def __init__(self):
        super().__init__()

        layout = QVBoxLayout()

        # Input fields
        self.value_input = QLineEdit()
        self.tolerance_input = QLineEdit()
        self.unit_combo = QComboBox()

        # Labels
        layout.addWidget(QLabel("Value:"))
        layout.addWidget(self.value_input)
        layout.addWidget(QLabel("Tolerance (%):"))
        layout.addWidget(self.tolerance_input)
        layout.addWidget(QLabel("Unit:"))
        layout.addWidget(self.unit_combo)

        # Component type selection
        self.component_type_combo = QComboBox()
        self.component_type_combo.addItems(["Capacitor", "Resistor", "Inductor"])
        self.component_type_combo.currentTextChanged.connect(self.update_unit_combo)

        layout.addWidget(QLabel("Component Type:"))
        layout.addWidget(self.component_type_combo)

        # Calculate button
        calculate_btn = QPushButton("Calculate")
        calculate_btn.clicked.connect(self.calculate)
        layout.addWidget(calculate_btn)

        self.result_label = QLabel()
        layout.addWidget(self.result_label)

        self.setLayout(layout)
        self.update_unit_combo("Capacitor")

    def update_unit_combo(self, component_type):
        self.unit_combo.clear()
        self.unit_combo.addItems(get_units(component_type))

    def calculate(self):
        try:
            value = float(self.value_input.text())
            tolerance = float(self.tolerance_input.text())
            min_value, max_value = calculate_with_tolerance(value, tolerance)
            unit = self.unit_combo.currentText()
            self.result_label.setText(f"Min: {min_value:.2f}{unit}, Max: {max_value:.2f}{unit}")
        except ValueError:
            QMessageBox.critical(self, "Error", "Please enter valid numeric values.")

# LCR Unit Converter Tab
class LCRUnitConverter(QWidget):
    def __init__(self):
        super().__init__()

        layout = QVBoxLayout()

        # Input fields
        self.value_input = QLineEdit()
        self.from_unit_combo = QComboBox()
        self.to_unit_combo = QComboBox()

        # Labels
        layout.addWidget(QLabel("Value:"))
        layout.addWidget(self.value_input)
        layout.addWidget(QLabel("From Unit:"))
        layout.addWidget(self.from_unit_combo)
        layout.addWidget(QLabel("To Unit:"))
        layout.addWidget(self.to_unit_combo)

        # Component type selection
        self.component_type_combo = QComboBox()
        self.component_type_combo.addItems(["Capacitor", "Resistor", "Inductor"])
        self.component_type_combo.currentTextChanged.connect(self.update_unit_combos)

        layout.addWidget(QLabel("Component Type:"))
        layout.addWidget(self.component_type_combo)

        # Convert button
        convert_btn = QPushButton("Convert")
        convert_btn.clicked.connect(self.convert)
        layout.addWidget(convert_btn)

        self.result_label = QLabel()
        layout.addWidget(self.result_label)

        self.setLayout(layout)
        self.update_unit_combos("Capacitor")

    def update_unit_combos(self, component_type):
        units = get_units(component_type)
        self.from_unit_combo.clear()
        self.from_unit_combo.addItems(units)
        self.to_unit_combo.clear()
        self.to_unit_combo.addItems(units)

    def convert(self):
        try:
            value = float(self.value_input.text())
            from_unit = self.from_unit_combo.currentText()
            to_unit = self.to_unit_combo.currentText()
            converted_value = convert_units(value, from_unit, to_unit)
            self.result_label.setText(f"Converted Value: {converted_value:.2e} {to_unit}")
        except ValueError:
            QMessageBox.critical(self, "Error", "Please enter a valid numeric value.")

# Data Viewer Tab
class DataViewer(QWidget):
    def __init__(self, data):
        super().__init__()

        layout = QVBoxLayout()
        self.table_view = QTableView()
        self.model = PandasModel(data)
        self.table_view.setModel(self.model)

        layout.addWidget(self.table_view)
        self.setLayout(layout)

# Main Window
class MainWindow(QMainWindow):
    def __init__(self, data):
        super().__init__()
        self.setWindowTitle("LCR Tolerance Calculator")

        tabs = QTabWidget()
        tabs.addTab(ComponentCalculator(), "Component Calculator")
        tabs.addTab(LCRUnitConverter(), "LCR Unit Converter")
        tabs.addTab(DataViewer(data), "LCR Data Viewer")

        self.setCentralWidget(tabs)

# Sample Data Creation
def create_data():
    """Create sample Pandas DataFrame for table display."""
    data = {
        "Component": ["Capacitor", "Resistor", "Inductor"],
        "Value": ["10uF", "1kΩ", "10mH"],
        "Tolerance": ["5%", "1%", "10%"]
    }
    return pd.DataFrame(data)

def main():
    app = QApplication(sys.argv)
    df = create_data()
    main_window = MainWindow(df)
    main_window.show()
    app.exec_()
    sys.exit()

if __name__ == "__main__":
    main()