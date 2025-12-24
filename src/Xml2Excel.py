from lxml import etree

def get_namespace(elem):
    return etree.QName(elem).namespace
def element_in_namespace(elem, target_ns):
    return get_namespace(elem) == target_ns

def flatten_xml(element, allowed_path, parent_path="", result=None, ignore_empty=False):
    if result is None:
        result = {}

    for child in element:
        name = etree.QName(child)
        tag = name.localname
        path = f"{parent_path}/{tag}" if parent_path else tag

        if len(child):
            flatten_xml(child, allowed_path, path, result, ignore_empty)
        else:
            if path in allowed_path and (ignore_empty is False or child.text):
                result[path] = (child.text or "").strip()
            else:
                print(f"Path {path} is not allowed path")

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
def extract_structure(parent, xml_node):

    q = etree.QName(xml_node)
    if q.localname:
        key = (q.namespace, q.localname)
        if key not in parent.children:
            parent.children[key] = Node(q.localname, q.namespace)
        for xml_node_child in xml_node:
            extract_structure(parent.children[key], xml_node_child)

def build_merged_tree(xml_files):
    root = Node("ROOT")

    for path in xml_files:
        tree = etree.parse(path)
        extract_structure(root, tree.getroot())

    return root

def parse_xml_file(path, allowed_path, ignore_empty=False):
    tree = etree.parse(path)
    root = tree.getroot()
    root_tag = etree.QName(root).localname
    rows = []
    for node in root.xpath("*"):
        node_tag = etree.QName(node).localname
        row = flatten_xml(node, allowed_path, f"{root_tag}/{node_tag}", None, ignore_empty)
        if row:
            row["_source_file"] = path  # полезно для отладки
            rows.append(row)
    return rows

import pandas as pd
def parse_multiple_xml(files, allowed_path, ignore_empty=False):
    all_rows = []
    for file in files:
        rows = parse_xml_file(file, allowed_path, ignore_empty)
        all_rows.extend(rows)
    return all_rows

def export_to_excel(df, output_path):
    df.to_excel(output_path, index=False)

import os
def collect_xml_files(root_folder: str) -> list[str]:
    xml_files = []
    for root, _, files in os.walk(root_folder):
        for filename in files:
            if filename.lower().endswith(".xml"):
                xml_files.append(os.path.join(root, filename))
    return xml_files

from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton,
    QListWidget, QFileDialog, QMessageBox,
    QTreeWidget, QTreeWidgetItem, QCheckBox
)
class XmlToExcelApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("XML → Excel")

        self.files = []

        layout = QVBoxLayout(self)

        self.btn_select = QPushButton("Выбрать XML-файлы")
        self.btn_select.clicked.connect(self.select_files)
        self.select_folder_btn = QPushButton("Выбрать папку с XML-файлами")
        self.select_folder_btn.clicked.connect(self.on_select_folder)

        self.ignore_empty_checkbox = QCheckBox("Игнорировать столбцы с пустыми ячейками")
        self.ignore_empty_checkbox.setChecked(True)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabel("Узлы XML")
        self.tree.itemChanged.connect(self.on_item_changed)
        self.tree.setExpandsOnDoubleClick(True)

        self.list_files = QListWidget()

        self.btn_export = QPushButton("Экспорт в Excel")
        self.btn_export.clicked.connect(self.export)

        layout.addWidget(self.btn_select)
        layout.addWidget(self.select_folder_btn)
        layout.addWidget(self.ignore_empty_checkbox)
        layout.addWidget(self.tree)
        layout.addWidget(self.btn_export)

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Выбор XML",
            "",
            "XML Files (*.xml)"
        )
        self.on_files_selected(files)

    def on_files_selected(self, files):
        self.files = files
        self.list_files.clear()
        self.list_files.addItems(files)
        merged_tree = build_merged_tree(files)
        populate_tree(self.tree, merged_tree)


    def on_select_folder(self):
        folder_path = QFileDialog.getExistingDirectory(
            self,
            "Выберите папку с XML-файлами",
            ""
        )
        if not folder_path:
            return

        xml_files = collect_xml_files(folder_path)

        if not xml_files:
            QMessageBox.warning(self, "Предупреждение", "В выбранной папке нет XML-файлов")
            return

        self.on_files_selected(xml_files)

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

        output, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить Excel",
            "",
            "Excel (*.xlsx)"
        )

        if not output:
            return

        try:
            checked_paths = set()
            for i in range(self.tree.topLevelItemCount()):
                top_item = self.tree.topLevelItem(i)
                collect_checked_paths(top_item, "", checked_paths)

            ignore_empty_fields = self.ignore_empty_checkbox.isChecked()
            rows = parse_multiple_xml(self.files, checked_paths, ignore_empty_fields)
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
        result = set()

    node = item.data(0, Qt.ItemDataRole.UserRole)
    path = f"{prefix}/{node.name}" if prefix else node.name

    if item.checkState(0) == Qt.CheckState.Checked and item.childCount() == 0:
        result.add(path)

    for i in range(item.childCount()):
        collect_checked_paths(item.child(i), path, result)

    return result

import sys
app = QApplication(sys.argv)
window = XmlToExcelApp()
window.show()
sys.exit(app.exec())