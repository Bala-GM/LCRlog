import sys
import pandas as pd
import PySimpleGUI as sg
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTableView, QVBoxLayout, QWidget, QTabWidget, 
    QGroupBox, QRadioButton, QComboBox, QLineEdit, QPushButton, QLabel, QMessageBox, QHBoxLayout
)
from PyQt5.QtCore import QAbstractTableModel, Qt

# Program 5: Database inspection interface_GUI/J0124-89P13
#def program_5():

# Model for displaying Pandas DataFrame in QTableView
class PandasModel(QAbstractTableModel):
    def __init__(self, data):
        super(PandasModel, self).__init__()
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid():
            if role == Qt.DisplayRole:
                return str(self._data.iloc[index.row(), index.column()])
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                return self._data.columns[section]
            if orientation == Qt.Vertical:
                return self._data.index[section]
        return None

# Calculators for component tolerances
def calculate_with_tolerance(value, lower_tolerance, upper_tolerance):
    lower_value = value - (value * lower_tolerance / 100)
    upper_value = value + (value * upper_tolerance / 100)
    return lower_value, upper_value

def capacitor_calculator(value, unit, lower_tolerance, upper_tolerance):
    units = {"F": 1, "mF": 1e-3, "µF": 1e-6, "uF": 1e-6, "nF": 1e-9, "pF": 1e-12}
    value_in_farads = value * units[unit]
    lower, upper = calculate_with_tolerance(value_in_farads, lower_tolerance, upper_tolerance)
    return value_in_farads, lower, upper

def resistor_calculator(value, unit, lower_tolerance, upper_tolerance):
    units = {"Ω": 1, "ohm": 1, "mΩ": 1e-3, "kΩ": 1e3, "MΩ": 1e6}
    value_in_ohms = value * units[unit]
    lower, upper = calculate_with_tolerance(value_in_ohms, lower_tolerance, upper_tolerance)
    return value_in_ohms, lower, upper

def inductor_calculator(value, unit, lower_tolerance, upper_tolerance):
    units = {"H": 1, "mH": 1e-3, "µH": 1e-6, "uH":1e-6, "nH": 1e-9, "kH": 1e3}
    value_in_henries = value * units[unit]
    lower, upper = calculate_with_tolerance(value_in_henries, lower_tolerance, upper_tolerance)
    return value_in_henries, lower, upper

# Unit conversion for LCR
def lcr_unit_conversion(value, from_unit, to_unit, unit_type):
    conversion_factors = {
        "Capacitance": {"F": 1, "mF": 1e-3, "µF": 1e-6, "uF":1e-6, "nF": 1e-9, "pF": 1e-12},
        "Resistance": {"Ω": 1, "ohm": 1, "mΩ": 1e-3, "kΩ": 1e3, "MΩ": 1e6},
        "Inductance": {"H": 1, "mH": 1e-3, "µH": 1e-6, "uH":1e-6, "nH": 1e-9, "kH": 1e3}
    }
    factor_from = conversion_factors[unit_type][from_unit]
    factor_to = conversion_factors[unit_type][to_unit]
    return value * (factor_from / factor_to)

# Main window class with tabs
class MainWindow(QMainWindow):
    def __init__(self, df):
        super().__init__()

        self.setWindowTitle("Component Calculator and LCR Help View")
        self.setGeometry(100, 100, 1280, 720)

        # Create central widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Create tabs
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)

        # Tab 1: Component Calculator
        self.component_calculator_tab = ComponentCalculator()
        self.tabs.addTab(self.component_calculator_tab, "Component Calculator")

        # Tab 2: LCR Unit Converter
        self.lcr_converter_tab = LCRUnitConverter()
        self.tabs.addTab(self.lcr_converter_tab, "LCR Unit Converter")

        # Tab 3: LCR Help View
        self.lcr_help_tab = QWidget()
        self.tabs.addTab(self.lcr_help_tab, "LCR Help View")
        lcr_layout = QVBoxLayout(self.lcr_help_tab)

        # Create QTableView for LCR data
        self.table_view = QTableView()
        model = PandasModel(df)
        self.table_view.setModel(model)
        lcr_layout.addWidget(self.table_view)

        # Show window maximized
        self.showMaximized()

# Component Calculator with tolerance calculations
class ComponentCalculator(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        # Component selection group
        self.component_group = QGroupBox("Select Component")
        self.capacitor_radio = QRadioButton("Capacitor")
        self.resistor_radio = QRadioButton("Resistor")
        self.inductor_radio = QRadioButton("Inductor")
        self.capacitor_radio.setChecked(True)  # Default to capacitor

        comp_layout = QHBoxLayout()
        comp_layout.addWidget(self.capacitor_radio)
        comp_layout.addWidget(self.resistor_radio)
        comp_layout.addWidget(self.inductor_radio)
        self.component_group.setLayout(comp_layout)
        layout.addWidget(self.component_group)

        # Value input and unit selection
        self.value_input = QLineEdit(self)
        self.value_input.setPlaceholderText("Enter value")
        layout.addWidget(QLabel("Enter Value:"))
        layout.addWidget(self.value_input)

        self.unit_combo = QComboBox(self)
        self.update_unit_combo("Capacitor")
        layout.addWidget(QLabel("Select Unit:"))
        layout.addWidget(self.unit_combo)

        # Lower and upper tolerance input
        self.lower_tolerance_input = QLineEdit(self)
        self.lower_tolerance_input.setPlaceholderText("Enter lower tolerance (%)")
        layout.addWidget(QLabel("Enter Lower Tolerance (%):"))
        layout.addWidget(self.lower_tolerance_input)

        self.upper_tolerance_input = QLineEdit(self)
        self.upper_tolerance_input.setPlaceholderText("Enter upper tolerance (%)")
        layout.addWidget(QLabel("Enter Upper Tolerance (%):"))
        layout.addWidget(self.upper_tolerance_input)

        # Calculate button
        calculate_btn = QPushButton("Calculate", self)
        calculate_btn.clicked.connect(self.calculate_result)
        layout.addWidget(calculate_btn)

        # Result label
        self.result_label = QLabel(self)
        layout.addWidget(self.result_label)

        # Update unit combo based on component selection
        self.capacitor_radio.toggled.connect(lambda: self.update_unit_combo("Capacitor"))
        self.resistor_radio.toggled.connect(lambda: self.update_unit_combo("Resistor"))
        self.inductor_radio.toggled.connect(lambda: self.update_unit_combo("Inductor"))

        self.setLayout(layout)

    def update_unit_combo(self, component_type):
        self.unit_combo.clear()
        if component_type == "Capacitor":
            self.unit_combo.addItems(["F", "mF", "µF", "uF", "nF", "pF"])
        elif component_type == "Resistor":
            self.unit_combo.addItems(["Ω", "ohm", "mΩ", "kΩ", "MΩ"])
        elif component_type == "Inductor":
            self.unit_combo.addItems(["H", "mH", "µH", "uH", "nH", "kH"])

    def calculate_result(self):
        try:
            value = float(self.value_input.text())
            unit = self.unit_combo.currentText()
            lower_tolerance = float(self.lower_tolerance_input.text())
            upper_tolerance = float(self.upper_tolerance_input.text())

            if self.capacitor_radio.isChecked():
                result_value, lower_value, upper_value = capacitor_calculator(value, unit, lower_tolerance, upper_tolerance)
            elif self.resistor_radio.isChecked():
                result_value, lower_value, upper_value = resistor_calculator(value, unit, lower_tolerance, upper_tolerance)
            elif self.inductor_radio.isChecked():
                result_value, lower_value, upper_value = inductor_calculator(value, unit, lower_tolerance, upper_tolerance)

            result_text = (f"Value: {result_value} {unit}\n"
                        f"Lower Tolerance Value: {lower_value} {unit}\n"
                        f"Upper Tolerance Value: {upper_value} {unit}")
            self.result_label.setText(result_text)
        except ValueError:
            QMessageBox.warning(self, "Input Error", "Please enter valid numeric values.")

# LCR Unit Converter class
class LCRUnitConverter(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        # Conversion type selection
        self.lcr_type_combo = QComboBox(self)
        self.lcr_type_combo.addItems(["Capacitance", "Resistance", "Inductance"])
        layout.addWidget(QLabel("Select Conversion Type:"))
        layout.addWidget(self.lcr_type_combo)

        # Value input and units selection
        self.lcr_value_input = QLineEdit(self)
        self.lcr_value_input.setPlaceholderText("Enter value")
        layout.addWidget(QLabel("Enter Value:"))
        layout.addWidget(self.lcr_value_input)

        self.from_unit_combo = QComboBox(self)
        self.to_unit_combo = QComboBox(self)
        self.update_conversion_units("Capacitance")  # Set default to Capacitance
        layout.addWidget(QLabel("From Unit:"))
        layout.addWidget(self.from_unit_combo)
        layout.addWidget(QLabel("To Unit:"))
        layout.addWidget(self.to_unit_combo)

        self.lcr_type_combo.currentTextChanged.connect(self.on_lcr_type_change)

        # Convert button
        convert_btn = QPushButton("Convert", self)
        convert_btn.clicked.connect(self.convert_lcr_units)
        layout.addWidget(convert_btn)

        self.conversion_result_label = QLabel(self)
        layout.addWidget(self.conversion_result_label)

        self.setLayout(layout)

    def update_conversion_units(self, unit_type):
        self.from_unit_combo.clear()
        self.to_unit_combo.clear()

        if unit_type == "Capacitance":
            units = ["F", "mF", "µF", "uF", "nF", "pF"]
        elif unit_type == "Resistance":
            units = ["Ω", "ohm", "mΩ", "kΩ", "MΩ"]
        elif unit_type == "Inductance":
            units = ["H", "mH", "µH", "uH", "nH", "kH"]

        self.from_unit_combo.addItems(units)
        self.to_unit_combo.addItems(units)

    def on_lcr_type_change(self, new_type):
        self.update_conversion_units(new_type)

    def convert_lcr_units(self):
        try:
            value = float(self.lcr_value_input.text())
            from_unit = self.from_unit_combo.currentText()
            to_unit = self.to_unit_combo.currentText()
            unit_type = self.lcr_type_combo.currentText()

            result = lcr_unit_conversion(value, from_unit, to_unit, unit_type)
            self.conversion_result_label.setText(f"Converted Value: {result} {to_unit}")
        except ValueError:
            QMessageBox.warning(self, "Input Error", "Please enter a valid numeric value.")

# Create sample data for LCR Help View
def create_data():
    columns = [
            'PartNumberName', 'VenderLotName', 'VENDERLOTPARTNUMBER.Basic settings_Part ShapeName',
            'VENDERLOTPARTNUMBER.Basic settings_PackageName', 'VENDERLOTPARTNUMBER.Basic settings_Barcode Label',
            'VENDERLOTPARTNUMBER.Specify direction_Direction', 'VENDERLOTPARTNUMBER.Specify other settings_LCR Check',
            'VENDERLOTPARTNUMBER.Specify other settings_LCR Check Parameter', 
            'VENDERLOTPARTNUMBER.Specify other settings_LCR Check Nominal Value', 
            'VENDERLOTPARTNUMBER.Specify other settings_LCR Check Nominal Value Unit',
            'VENDERLOTPARTNUMBER.Specify other settings_LCR Check Tolerance',
            'VENDERLOTPARTNUMBER.Specify other settings_LCR Check Frequency',
            'VENDERLOTPARTNUMBER.Specify other settings_LCR Check Frequency Unit',
            'VENDERLOTPARTNUMBER.Specify other settings_LCR Check Current'
        ]

    data = [
        ['DumpC1 -', 'Vender1', 'Capacitor', '0804P', '-', 0, 1, 1, 10, 0, 20, 200, 3, 1],
        ['DumpC2 m', 'Vender1', 'Capacitor', '0804P', 'm', 0, 1, 1, 10, -3, 20, 200, 3, 1],
        ['DumpC3 u', 'Vender1', 'Capacitor', '0804P', 'u', 0, 1, 1, 10, -6, 20, 200, 3, 1],
        ['DumpC4 n', 'Vender1', 'Capacitor', '0804P', 'n', 0, 1, 1, 10, -9, 20, 200, 3, 1],
        ['DumpC5 p', 'Vender1', 'Capacitor', '0804P', 'p', 0, 1, 1, 10, -12, 20, 200, 3, 1],
        ['-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-'],
        ['DumpI1 k', 'Vender1', 'Inductor', '0804P', 'k', 0, 1, 0, 10, 3, 20, 200, 3, 1],
        ['DumpI2 -', 'Vender1', 'Inductor', '0804P', '-', 0, 1, 0, 10, 0, 20, 200, 3, 1],
        ['DumpI3 m', 'Vender1', 'Inductor', '0804P', 'm', 0, 1, 0, 10, -3, 20, 200, 3, 1],
        ['DumpI4 u', 'Vender1', 'Inductor', '0804P', 'u', 0, 1, 0, 10, -6, 20, 200, 3, 1],
        ['DumpI5 n', 'Vender1', 'Inductor', '0804P', 'n', 0, 1, 0, 10, -9, 20, 200, 3, 1],
        ['-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-'],
        ['DumpR1 M', 'Vender1', 'Resistor', '0804P', 'M', 0, 1, 2, 10, 6, 5, 40, 0, 1],
        ['DumpR2 k', 'Vender1', 'Resistor', '0804P', 'k', 0, 1, 2, 10, 3, 5, 40, 0, 1],
        ['DumpR3 -', 'Vender1', 'Resistor', '0804P', '-', 0, 1, 2, 10, 0, 5, 40, 0, 1],
        ['DumpR4 m', 'Vender1', 'Resistor', '0804P', 'm', 0, 1, 2, 10, -3, 5, 40, 0, 1]
    ]

    return pd.DataFrame(data, columns=columns)

#if __name__ == '__main__':
app = QApplication(sys.argv)
df = create_data()
main_window = MainWindow(df)
main_window.show()
#sys.exit(app.exec_())
app.exec_()
sg.Popup('Copyright (C) <2024>  <BALA GANESH>', 'Thanks For using LCR Tolerance Hint')
sys.exit()