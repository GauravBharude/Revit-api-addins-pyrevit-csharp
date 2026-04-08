# -*- coding: utf-8 -*-

from System.Windows.Controls import CheckBox
from pyrevit import revit, forms, DB
from pyrevit.framework import wpf
import os

doc = revit.doc


class ExportUI(forms.WPFWindow):
    def __init__(self):
        xaml_path = os.path.join(os.path.dirname(__file__), 'ui.xaml')
        wpf.LoadComponent(self, xaml_path)

        self.export_map = {}
        self.export_path = None
        self.view_name = None

        # ---------------- Worksets ----------------
        self.worksets = [
            ws.Name for ws in
            DB.FilteredWorksetCollector(doc)
            .OfKind(DB.WorksetKind.UserWorkset)
        ]

        for ws in self.worksets:
            cb = CheckBox()
            cb.Content = ws
            self.WorksetPanel.Children.Add(cb)

        # ---------------- 3D Views ----------------
        self.views = [
            v for v in
            DB.FilteredElementCollector(doc)
            .OfClass(DB.View3D)
            if not v.IsTemplate
        ]

        for v in self.views:
            self.ViewCombo.Items.Add(v)

        if self.views:
            self.ViewCombo.SelectedIndex = 0

    # ---------- UI EVENTS ----------

    def browse_folder(self, sender, args):
        folder = forms.pick_folder()
        if folder:
            self.PathBox.Text = folder

    def add_nwc(self, sender, args):
        nwc_name = self.NwcNameBox.Text.strip()

        if not nwc_name:
            forms.alert("Please enter an NWC name")
            return

        selected_ws = [
            cb.Content for cb in self.WorksetPanel.Children
            if cb.IsChecked is True
        ]

        if not selected_ws:
            forms.alert("Select at least one workset")
            return

        # Prevent duplicate NWC names
        if nwc_name in self.export_map:
            forms.alert("NWC name already exists")
            return

        # Save data
        self.export_map[nwc_name] = selected_ws

        # ---- SHOW IN UI SUMMARY ----
        summary = "{}  →  {}".format(
            nwc_name,
            ", ".join(selected_ws)
        )
        self.NwcSummaryList.Items.Add(summary)

        # Reset inputs
        self.NwcNameBox.Text = ""
        for cb in self.WorksetPanel.Children:
            cb.IsChecked = False

        forms.toast("NWC Added: {}".format(nwc_name))

    def submit(self, sender, args):
        if not self.PathBox.Text:
            forms.alert("Select an export folder")
            return

        if not self.export_map:
            forms.alert("Add at least one NWC definition")
            return

        self.export_path = self.PathBox.Text
        self.view_name = self.ViewCombo.SelectedItem.Name

        self.Close()
