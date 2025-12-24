from lxml import etree

def get_namespace(elem):
    return etree.QName(elem).namespace
def element_in_namespace(elem, target_ns):
    return get_namespace(elem) == target_ns

def flatten_xml(element, parent_path="", result=None):
    if result is None:
        result = {}

    for child in element:
        name = etree.QName(child)
        tag = name.localname
        path = f"{parent_path}/{tag}" if parent_path else tag

        if len(child):
            flatten_xml(child, path, result)
        else:
            result[path] = (child.text or "").strip()

    return result

class Node:
    def __init__(self, name, namespace=None):
        self.name = name
        self.namespace = namespace
        self.children = {}

    def key(self):
        return self.namespace, self.name
# -------------------------------
# XML utils
# -------------------------------
def extract_structure(elem, node):
    for child in elem:
        q = etree.QName(child)
        key = (q.namespace, q.localname)

        if key not in node.children:
            node.children[key] = Node(q.localname, q.namespace)

        extract_structure(child, node.children[key])

def build_merged_tree(xml_files):
    root = Node("ROOT")

    for path in xml_files:
        tree = etree.parse(path)
        extract_structure(tree.getroot(), root)

    return root

def collect_namespaces(root):
    namespaces = set()
    for elem in root.iter():
        ns = etree.QName(elem).namespace
        if ns:
            namespaces.add(ns)
    return namespaces

def parse_xml_file(path, row_xpath):
    tree = etree.parse(path)
    root = tree.getroot()
    rows = []
    for node in root.xpath(row_xpath):
        row = flatten_xml(node)
        row["_source_file"] = path  # полезно для отладки
        rows.append(row)

    return rows

import pandas as pd
def parse_multiple_xml(files, row_xpath):
    all_rows = []
    for file in files:
        rows = parse_xml_file(file, row_xpath)
        all_rows.extend(rows)
    return all_rows

def export_to_excel(df, output_path):
    df.to_excel(output_path, index=False)

from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton,
    QListWidget, QFileDialog, QLineEdit, QLabel, QMessageBox,
    QTreeWidget, QTreeWidgetItem
)
class XmlToExcelApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("XML → Excel")

        self.files = []

        layout = QVBoxLayout(self)

        self.btn_select = QPushButton("Выбрать XML файлы")
        self.btn_select.clicked.connect(self.select_files)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabel("Узлы XML")
        self.tree.itemChanged.connect(self.on_item_changed)
        self.tree.setExpandsOnDoubleClick(True)

        self.list_files = QListWidget()

        self.xpath_label = QLabel("XPath узла-строки:")
        self.xpath_input = QLineEdit(".//*")

        self.btn_export = QPushButton("Экспорт в Excel")
        self.btn_export.clicked.connect(self.export)

        layout.addWidget(self.btn_select)
        layout.addWidget(self.tree)
        layout.addWidget(self.xpath_label)
        layout.addWidget(self.xpath_input)
        layout.addWidget(self.btn_export)

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Выбор XML",
            "",
            "XML Files (*.xml)"
        )
        self.files = files
        self.list_files.clear()
        self.list_files.addItems(files)
        merged_tree = build_merged_tree(files)
        populate_tree(self.tree, merged_tree)

    def on_item_changed(self, item, column):
        # Игнорируем изменения без чекбокса
        if column != 0:
            return

        state = item.checkState(0)
        self.tree.blockSignals(True)
        # 🔹 обновляем всех детей
        update_children(item, state)
        # 🔹 обновляем родителей
        update_parent(item.parent())
        self.tree.blockSignals(False)

    def export(self):
        if not self.files:
            QMessageBox.warning(self, "Ошибка", "Файлы не выбраны")
            return

        xpath = self.xpath_input.text().strip()
        if not xpath:
            QMessageBox.warning(self, "Ошибка", "XPath не задан")
            return

        output, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить Excel",
            "",
            "Excel (*.xlsx)"
        )

        if not output:
            return

        try:
            rows = parse_multiple_xml(self.files, xpath)
            df = pd.DataFrame(rows)

            export_to_excel(df, output)
            QMessageBox.information(self, "Готово", "Экспорт завершён")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))

from PyQt6.QtCore import Qt

def update_children(item, state):
    for i in range(item.childCount()):
        child = item.child(i)
        child.setCheckState(0, state)
        update_children(child, state)

def update_parent(item):
    if item is None:
        return

    checked = 0
    unchecked = 0

    for i in range(item.childCount()):
        child_state = item.child(i).checkState(0)
        if child_state == Qt.CheckState.Checked:
            checked += 1
        elif child_state == Qt.CheckState.Unchecked:
            unchecked += 1
        else:
            # PartiallyChecked
            checked += 1
            unchecked += 1

    if checked == item.childCount() and unchecked == 0:
        item.setCheckState(0, Qt.CheckState.Checked)
    elif unchecked == item.childCount() and checked == 0:
        item.setCheckState(0, Qt.CheckState.Unchecked)
    else:
        item.setCheckState(0, Qt.CheckState.PartiallyChecked)

    # рекурсивно вверх
    update_parent(item.parent())

def populate_tree(widget, node):
    widget.blockSignals(True)
    widget.clear()
    for child in node.children.values():
        add_item(widget, child)
    widget.blockSignals(False)

def add_item(parent, node):
    label = f'{node.name}'
    if node.namespace:
        label = f'{{{node.namespace}}}{node.name}'
    item = QTreeWidgetItem([label])
    item.setCheckState(0, Qt.CheckState.Unchecked)
    item.setData(0, Qt.ItemDataRole.UserRole, node)
    if isinstance(parent, QTreeWidgetItem):
        parent.addChild(item)
    else:
        parent.addTopLevelItem(item)
    for child in node.children.values():
        add_item(item, child)

def collect_checked_paths(item, prefix="", result=None):
    if result is None:
        result = []

    node = item.data(0, Qt.ItemDataRole.UserRole)
    path = f"{prefix}.{node.name}" if prefix else node.name

    if item.checkState(0) == Qt.CheckState.Checked:
        result.append(path)

    for i in range(item.childCount()):
        collect_checked_paths(item.child(i), path, result)

    return result

import sys
app = QApplication(sys.argv)
window = XmlToExcelApp()
window.show()
sys.exit(app.exec())