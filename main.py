import flet as ft
from routes.router import Router

def main(page: ft.Page):
    router = Router(page)
    router.setup_main_route()

if __name__ == "__main__":
    ft.app(target=main)
