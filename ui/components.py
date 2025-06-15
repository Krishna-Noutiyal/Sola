import flet as ft
from config import ColorScheme
from scripts import ExcelProcessor # type: ignore
import os

class MainView:
    def __init__(self, page: ft.Page):
        self.page = page
        self.excel_processor = ExcelProcessor()
        self.selected_files = []
        self.output_path = ""

        # UI Components
        self.file_picker = ft.FilePicker(on_result=self.on_files_selected)
        self.output_picker = ft.FilePicker(on_result=self.on_output_selected)
        self.page.overlay.extend([self.file_picker, self.output_picker])

        self.selected_file_text = ft.Text(
            "No File Selected",
            color=ColorScheme.TEXT_SECONDARY,
            size=14
        )

        self.output_path_text = ft.Text(
            "No Form-16 Selected",
            color=ColorScheme.TEXT_SECONDARY,
            size=14
        )

        self.status_text = ft.Text(
            "",
            color=ColorScheme.TEXT_SECONDARY,
            size=14
        )

    def on_files_selected(self, e: ft.FilePickerResultEvent):
        if e.files:
            self.selected_files = [file.path for file in e.files]
            file_names = [os.path.basename(path) for path in self.selected_files]
            self.selected_file_text.value = f"ITR Format: {', '.join(file_names)}"
            self.selected_file_text.color = ColorScheme.SUCCESS
        else:
            self.selected_files = []
            self.selected_file_text.value = "No File Selected"
            self.selected_file_text.color = ColorScheme.TEXT_SECONDARY
        self.page.update()

    def on_output_selected(self, e: ft.FilePickerResultEvent):
        if e.path:
            self.output_path = e.path
            self.output_path_text.value = f"Form-16: {os.path.basename(self.output_path)}"
            self.output_path_text.color = ColorScheme.SUCCESS
        else:
            self.output_path = ""
            self.output_path_text.value = "No Form-16 Selected"
            self.output_path_text.color = ColorScheme.TEXT_SECONDARY
        self.page.update()

    def on_submit_clicked(self, e):
        if not self.selected_files:
            self.show_status("Please Select ITR Format !", ColorScheme.ERROR)
            return

        if not self.output_path:
            self.show_status("Please Select Form-16 !", ColorScheme.ERROR)
            return

        try:
            self.show_status("Processing File...", ColorScheme.PRIMARY)

            # Call the ExcelProcessor to create Form-16
            create_Excel = self.excel_processor.create_form_16(
                itr_format=self.selected_files[0],  # Assuming single file selection
                form_16=self.output_path,
            )

            if create_Excel:
                self.show_status("Form-16 Filled Successfully !", ColorScheme.SUCCESS)
            else:
                self.show_status("Error Processing File !", ColorScheme.ERROR)
        except Exception as ex:
            self.show_status(f"Error: {str(ex)}", ColorScheme.ERROR)

    def show_status(self, message: str, color: str):
        self.status_text.value = message
        self.status_text.color = color
        self.status_text.weight = ft.FontWeight.BOLD
        self.page.update()

    def build(self):
        return ft.Container(
            # width= self.page.width,
            # height= self.page.height,
            content=ft.Column([
                # Title
                ft.Container(
                    content=ft.Text(
                        "Sola : Form-16 Generator",
                        size=32,
                        weight=ft.FontWeight.BOLD,
                        color=ColorScheme.PRIMARY
                    ),
                    margin=ft.margin.only(bottom=10)
                ),
                
                # Description
                ft.Container(
                    content=ft.Text(
                        "Hello, Sola is a Form-16 Filler, Use the ITR format of user to fill the desired Form-16 (xlsx) file of the respective user.\n",
                        size=16,
                        color=ColorScheme.TEXT_SECONDARY
                    ),
                    margin=ft.margin.only(bottom=30)
                ),
                # File Selection Section
                ft.Container(
                    content=ft.Column([
                        ft.Text(
                            "Select ITR Format:",
                            size=18,
                            weight=ft.FontWeight.W_500,
                            color=ColorScheme.TEXT_PRIMARY
                        ),
                        ft.Container(
                            content=ft.Row([
                                ft.ElevatedButton(
                                    "ITR Format (PIC)",
                                    icon=ft.Icons.FOLDER_OPEN,
                                    on_click=lambda _: self.file_picker.pick_files(
                                        allow_multiple=False,
                                        allowed_extensions=["xlsx"]
                                    ),
                                    bgcolor=ColorScheme.PRIMARY,
                                    color=ft.Colors.WHITE,
                                    width=200,
                                    height=50,
                                    style=ft.ButtonStyle(
                                        text_style=ft.TextStyle(size=16,weight=ft.FontWeight.BOLD)  # Increased text size
                                        
                                    ),
                                )
                            ]),
                            margin=ft.margin.only(top=5, bottom=10)
                        ),
                        self.selected_file_text
                    ]),
                    padding=20,
                    border=ft.border.all(1, ColorScheme.BORDER),
                    border_radius=8,
                    bgcolor=ColorScheme.SURFACE,
                    margin=ft.margin.only(bottom=20)
                ),
                
                # Select Form-15 Selection Section
                ft.Container(
                    content=ft.Column([
                        ft.Text(
                            "Select Form-16:",
                            size=18,
                            weight=ft.FontWeight.W_500,
                            color=ColorScheme.TEXT_PRIMARY
                        ),
                        ft.Container(
                            content=ft.Row([
                                ft.ElevatedButton(
                                    "Form-16",
                                    icon=ft.Icons.SAVE,
                                    on_click=lambda _: self.output_picker.save_file(
                                        file_name="Form-16.xlsx",
                                        allowed_extensions=["xlsx"]
                                    ),
                                    bgcolor=ColorScheme.SECONDARY,
                                    color=ColorScheme.TEXT_PRIMARY,
                                    width=200,
                                    height=50,
                                    style=ft.ButtonStyle(
                                        text_style=ft.TextStyle(size=16,weight=ft.FontWeight.BOLD)  # Increased text size
                                        
                                    ),
                                )
                            ]),
                            margin=ft.margin.only(top=5, bottom=10)
                        ),
                        self.output_path_text
                    ]),
                    padding=20,
                    border=ft.border.all(1, ColorScheme.BORDER),
                    border_radius=8,
                    bgcolor=ColorScheme.SURFACE,
                    margin=ft.margin.only(bottom=30)
                ),
                
                # Submit Button
                ft.Container(
                    content=ft.ElevatedButton(
                        "Submit",
                        icon=ft.Icons.PLAY_ARROW,
                        on_click=self.on_submit_clicked,
                        bgcolor=ColorScheme.SUCCESS,
                        color=ft.Colors.WHITE,
                        width=200,
                        height=50,
                        style=ft.ButtonStyle(
                            text_style=ft.TextStyle(size=16, weight=ft.FontWeight.BOLD)  # Increased text size
                        )
                    ),
                    alignment=ft.alignment.center,
                    margin=ft.margin.only(bottom=10)
                ),
                
                # Status Text
                ft.Container(
                    content=self.status_text,
                    alignment=ft.alignment.center,
                )
            ]),
            bgcolor=ColorScheme.BACKGROUND,
            padding=50,
            expand=True,
            border_radius=15
        )
