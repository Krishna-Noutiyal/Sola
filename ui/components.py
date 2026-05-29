import flet as ft
from config import ColorScheme
from scripts import ExcelProcessor  # type: ignore
import os


class MainView:
    def __init__(self, page: ft.Page):
        self.page = page
        self.excel_processor = ExcelProcessor()
        self.selected_files = ""
        self.output_path = ""
        self.file_path = ""

        self.selected_file_text = ft.Text(
            "No File Selected", color=ColorScheme.TEXT_SECONDARY, size=14
        )

        self.output_path_text = ft.Text(
            "No Form-16 Selected", color=ColorScheme.TEXT_SECONDARY, size=14
        )

        self.status_text = ft.Text("", color=ColorScheme.TEXT_SECONDARY, size=14)

    async def pick_file(self, e: ft.Event[ft.Button]):
        files = await ft.FilePicker().pick_files(
            allow_multiple=False,
            allowed_extensions=["xlsx"],
        )
        if files:
            self.selected_file = files[0]
            self.file_path = self.selected_file.path
            file_name = self.selected_file.name
            self.selected_file_text.value = f"ITR Format: {file_name}"
            self.selected_file_text.color = ColorScheme.SUCCESS
        else:
            self.selected_files = ""
            self.file_path = ""
            self.selected_file_text.value = "No File Selected"
            self.selected_file_text.color = ColorScheme.TEXT_SECONDARY
        self.page.update()

    async def pick_output(self, e: ft.Event[ft.Button]):
        file_path = await ft.FilePicker().save_file(
            file_name="Form-16.xlsx",
            allowed_extensions=["xlsx"],
        )
        if file_path:
            self.output_path = file_path
            self.output_path_text.value = f"Form-16: {os.path.basename(file_path)}"
            self.output_path_text.color = ColorScheme.SUCCESS
        else:
            self.output_path = ""
            self.output_path_text.value = "No Form-16 Selected"
            self.output_path_text.color = ColorScheme.TEXT_SECONDARY
        self.page.update()

    def on_submit_clicked(self, e):
        if not self.selected_file:
            self.show_status("Please Select ITR Format !", ColorScheme.ERROR)
            return

        if not self.output_path:
            self.show_status("Please Select Form-16 !", ColorScheme.ERROR)
            return

        try:
            self.show_status("Processing File...", ColorScheme.PRIMARY)

            # Call the ExcelProcessor to create Form-16
            create_Excel = self.excel_processor.create_form_16(
                itr_format=self.file_path or "",
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
            content=ft.Column(
                [
                    # Title
                    ft.Container(
                        content=ft.Text(
                            "Sola : Form-16 Generator",
                            size=32,
                            weight=ft.FontWeight.BOLD,
                            color=ColorScheme.PRIMARY,
                        ),
                        margin=ft.Margin(bottom=30),
                    ),
                    # Description
                    ft.Container(
                        content=ft.Text(
                            "Hello, Sola is a Form-16 Filler, Use the ITR format of user to fill the desired Form-16 (xlsx) file of the respective user.\n",
                            size=16,
                            color=ColorScheme.TEXT_SECONDARY,
                        ),
                        margin=ft.Margin(bottom=30),
                    ),
                    # File Selection Section
                    ft.Container(
                        content=ft.Column(
                            [
                                ft.Text(
                                    "Select ITR Format:",
                                    size=18,
                                    weight=ft.FontWeight.W_500,
                                    color=ColorScheme.TEXT_PRIMARY,
                                ),
                                ft.Container(
                                    content=ft.Row(
                                        [
                                            ft.ElevatedButton(
                                                "ITR Format (PIC)",
                                                icon=ft.Icons.FOLDER_OPEN,
                                                on_click=self.pick_file,
                                                bgcolor=ColorScheme.PRIMARY,
                                                color=ft.Colors.WHITE,
                                                width=200,
                                                height=50,
                                                style=ft.ButtonStyle(
                                                    text_style=ft.TextStyle(
                                                        size=16,
                                                        weight=ft.FontWeight.BOLD,
                                                    )
                                                ),
                                            )
                                        ]
                                    ),
                                    margin=ft.Margin(top=5, bottom=10),
                                ),
                                self.selected_file_text,
                            ]
                        ),
                        padding=20,  # all(1, ColorScheme.BORDER)
                        border=ft.Border.all(1, ColorScheme.BORDER),
                        border_radius=8,
                        bgcolor=ColorScheme.SURFACE,
                        margin=ft.Margin(bottom=20),
                    ),
                    # Select Form-15 Selection Section
                    ft.Container(
                        content=ft.Column(
                            [
                                ft.Text(
                                    "Select Form-16:",
                                    size=18,
                                    weight=ft.FontWeight.W_500,
                                    color=ColorScheme.TEXT_PRIMARY,
                                ),
                                ft.Container(
                                    content=ft.Row(
                                        [
                                            ft.ElevatedButton(
                                                "Form-16",
                                                icon=ft.Icons.SAVE,
                                                on_click=self.pick_output,
                                                bgcolor=ColorScheme.SECONDARY,
                                                color=ColorScheme.TEXT_PRIMARY,
                                                width=200,
                                                height=50,
                                                style=ft.ButtonStyle(
                                                    text_style=ft.TextStyle(
                                                        size=16,
                                                        weight=ft.FontWeight.BOLD,
                                                    )
                                                ),
                                            )
                                        ]
                                    ),
                                    margin=ft.Margin(top=5, bottom=10),
                                ),
                                self.output_path_text,
                            ]
                        ),
                        padding=20,
                        border=ft.Border.all(1, ColorScheme.BORDER),
                        border_radius=8,
                        bgcolor=ColorScheme.SURFACE,
                        margin=ft.Margin(bottom=30),
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
                                text_style=ft.TextStyle(
                                    size=16, weight=ft.FontWeight.BOLD
                                )  # Increased text size
                            ),
                        ),
                        alignment=ft.Alignment.CENTER,
                        margin=ft.Margin(bottom=10),
                    ),
                    # Status Text
                    ft.Container(
                        content=self.status_text,
                        alignment=ft.Alignment.CENTER,
                    ),
                ]
            ),
            bgcolor=ColorScheme.BACKGROUND,
            padding=50,
            expand=True,
            border_radius=15,
        )
