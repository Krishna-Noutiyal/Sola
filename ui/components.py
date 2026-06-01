import flet as ft
from config import ColorScheme
from scripts import ExcelProcessor  # type: ignore
import os
import asyncio


# TODO Create a Toggle button for Retired Emp.
class MainView:
    def __init__(self, page: ft.Page):
        self.page = page
        self.is_retired = False
        self.selected_file = ""
        self.output_path = ""
        self.file_path = ""
        self.is_processing = False

        self.selected_file_text = ft.Text(
            "No File Selected", color=ColorScheme.TEXT_SECONDARY, size=14
        )

        self.output_path_text = ft.Text(
            "No Form-16 Selected", color=ColorScheme.TEXT_SECONDARY, size=14
        )

        self.status_text = ft.Text("", color=ColorScheme.TEXT_SECONDARY, size=14)

        self.progress_bar = ft.ProgressBar(
            width=300,
            color=ColorScheme.PRIMARY,
            bgcolor=ColorScheme.SURFACE,
            visible=False,
        )

        self.retired_toggle = ft.Switch(
            label="Retired Employee",
            label_position=ft.LabelPosition.RIGHT,
            on_change=self.on_retired_toggle_changed,
        )

        self.submit_button = ft.ElevatedButton(
            "Submit",
            icon=ft.Icons.PLAY_ARROW,
            on_click=self.on_submit_clicked,
            bgcolor=ColorScheme.SUCCESS,
            color=ft.Colors.WHITE,
            width=200,
            height=50,
            style=ft.ButtonStyle(
                text_style=ft.TextStyle(size=16, weight=ft.FontWeight.BOLD)
            ),
        )

    async def pick_file(self, e: ft.Event[ft.Button]):
        files = await ft.FilePicker().pick_files(
            allow_multiple=False,
            allowed_extensions=["xlsx"],
        )
        if files:
            self.selected_file = files[0]
            self.file_path = self.selected_file.path
            file_name = self.selected_file.name
            self.selected_file_text.value = f"📄 ITR Format: {file_name}"
            self.selected_file_text.color = ColorScheme.SUCCESS
        else:
            self.selected_files = ""
            self.file_path = ""
            self.selected_file_text.value = "No File Selected"
            self.selected_file_text.color = ColorScheme.TEXT_SECONDARY
        self._update_submit_button()
        self.page.update()

    async def pick_output(self, e: ft.Event[ft.Button]):
        file_path = await ft.FilePicker().save_file(
            file_name="Form-16.xlsx",
            allowed_extensions=["xlsx"],
        )
        if file_path:
            self.output_path = file_path
            self.output_path_text.value = f"📁 Form-16: {os.path.basename(file_path)}"
            self.output_path_text.color = ColorScheme.SUCCESS
        else:
            self.output_path = ""
            self.output_path_text.value = "No Form-16 Selected"
            self.output_path_text.color = ColorScheme.TEXT_SECONDARY
        self._update_submit_button()
        self.page.update()

    def _update_submit_button(self):
        has_file = bool(self.selected_file)
        has_output = bool(self.output_path)
        self.submit_button.disabled = not (has_file and has_output)
        if not has_file and not has_output:
            self.submit_button.bgcolor = ColorScheme.SURFACE
        else:
            self.submit_button.bgcolor = ColorScheme.SUCCESS

    async def on_submit_clicked(self, e):
        if not self.selected_file:
            self.show_status("⚠️ Please Select ITR Format !", ColorScheme.ERROR)
            return

        if not self.output_path:
            self.show_status("⚠️ Please Select Form-16 !", ColorScheme.ERROR)
            return

        if self.is_processing:
            return

        try:
            self.is_processing = True
            self.submit_button.disabled = True
            self.submit_button.content = "Processing..."
            self.submit_button.bgcolor = ColorScheme.PRIMARY
            self.progress_bar.visible = True
            self.show_status("⏳ Processing File...", ColorScheme.PRIMARY)
            self.page.update()

            # Ensure minimum processing time for UX feedback
            await asyncio.sleep(0.5)

            # Call the ExcelProcessor to create Form-16
            excel_processor = ExcelProcessor(self.is_retired)
            create_Excel = excel_processor.create_form_16(
                itr_format=self.file_path or "",
                form_16=self.output_path,
            )

            self.is_processing = False
            self.progress_bar.visible = False

            if create_Excel:
                self.show_status(
                    "✅ Form-16 Filled Successfully !", ColorScheme.SUCCESS
                )
            else:
                self.show_status("❌ Error Processing File !", ColorScheme.ERROR)
        except Exception as ex:
            self.is_processing = False
            self.progress_bar.visible = False
            self.show_status(f"❌ Error: {str(ex)}", ColorScheme.ERROR)
        finally:
            self.submit_button.disabled = False
            self.submit_button.content = "Submit"
            self._update_submit_button()
            self.page.update()

    def show_status(self, message: str, color: str):
        self.status_text.value = message
        self.status_text.color = color
        self.status_text.weight = ft.FontWeight.BOLD
        self.page.update()

    def on_retired_toggle_changed(self, e):
        self.is_retired = e.control.value
        self.page.update()

    def build(self):
        return ft.Container(
            content=ft.Column(
                [
                    # Title
                    ft.Container(
                        content=ft.Row(
                            [
                                ft.Image(
                                    src=r"assets\icon.png",
                                    width=48,
                                    height=48,
                                    fit=ft.BoxFit.CONTAIN,
                                ),
                                ft.Text(
                                    "Sola : Form-16 Generator",
                                    size=32,
                                    weight=ft.FontWeight.BOLD,
                                    color=ColorScheme.PRIMARY,
                                ),
                            ],
                            alignment=ft.MainAxisAlignment.CENTER,
                        ),
                        margin=ft.Margin(bottom=20),
                    ),
                    # Description
                    ft.Container(
                        content=ft.Text(
                            "Fill Form-16 using ITR format data of the employee",
                            size=16,
                            color=ColorScheme.TEXT_SECONDARY,
                            text_align=ft.TextAlign.CENTER,
                        ),
                        margin=ft.Margin(bottom=30),
                    ),
                    # File Selection Section
                    ft.Container(
                        content=ft.Column(
                            [
                                ft.Row(
                                    [
                                        ft.Icon(
                                            ft.Icons.DESCRIPTION,
                                            size=20,
                                            color=ColorScheme.TEXT_SECONDARY,
                                        ),
                                        ft.Text(
                                            "Select ITR Format:",
                                            size=18,
                                            weight=ft.FontWeight.W_500,
                                            color=ColorScheme.TEXT_PRIMARY,
                                        ),
                                    ],
                                    spacing=10,
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
                                            ),
                                            ft.Container(
                                                content=self.retired_toggle,
                                                margin=ft.Margin(left=20),
                                            ),
                                        ]
                                    ),
                                    margin=ft.Margin(top=5, bottom=10),
                                ),
                                self.selected_file_text,
                            ]
                        ),
                        padding=20,
                        border=ft.Border.all(1, ColorScheme.BORDER),
                        border_radius=8,
                        bgcolor=ColorScheme.SURFACE,
                        margin=ft.Margin(bottom=20),
                    ),
                    # Select Form-16 Selection Section
                    ft.Container(
                        content=ft.Column(
                            [
                                ft.Row(
                                    [
                                        ft.Icon(
                                            ft.Icons.SAVE,
                                            size=20,
                                            color=ColorScheme.TEXT_SECONDARY,
                                        ),
                                        ft.Text(
                                            "Select Form-16:",
                                            size=18,
                                            weight=ft.FontWeight.W_500,
                                            color=ColorScheme.TEXT_PRIMARY,
                                        ),
                                    ],
                                    spacing=10,
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
                    # Progress Bar
                    ft.Container(
                        content=self.progress_bar,
                        alignment=ft.Alignment.CENTER,
                        margin=ft.Margin(bottom=15),
                    ),
                    # Submit Button
                    ft.Container(
                        content=self.submit_button,
                        alignment=ft.Alignment.CENTER,
                        margin=ft.Margin(bottom=15),
                    ),
                    # Status Text
                    ft.Container(
                        content=self.status_text,
                        alignment=ft.Alignment.CENTER,
                    ),
                ],
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                spacing=10,
            ),
            bgcolor=ColorScheme.BACKGROUND,
            padding=50,
            expand=True,
            border_radius=15,
        )
