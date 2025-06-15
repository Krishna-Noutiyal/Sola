import flet as ft
from ui import MainView
import toml
import os

class Router:
    def __init__(self, page: ft.Page):
        self.page = page
        self.setup_page()
    
    def setup_page(self):
        # Get path to myproject.toml (one directory above)
        toml_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "myproject.toml")
        with open(toml_path, "r") as f:
            data = toml.load(f)
            
        project = data.get("project", {})
        name = project.get("name", "")
        version = project.get("version", "")
        
        self.page.title = f"{name} v{version}"
        self.page.theme_mode = ft.ThemeMode.DARK
        self.page.auto_scroll = True
        self.page.window.height = 800
        # self.page.padding = 20
        # self.page.expand = True
    
    def setup_main_route(self):
        main_view = MainView(self.page)
        self.page.add(main_view.build())
        self.page.update()