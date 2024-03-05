from collections import OrderedDict
from pathlib import Path

import flet as ft
from loguru import logger
from pytube import Playlist

from utils import XLSXFile


logger.add("error.log", format="{time} {level} {message}", level="ERROR")


@logger.catch
def main(page: ft.Page):

    def on_btn_export_click(e: ft.ControlEvent):

        try:
            if not text_field.value:
                return

            btn_export.disabled = True
            progress_bar.value = 0
            page.update()

            link = text_field.value
            playlist = Playlist(link)
            videos = list(playlist.videos)
            text_label.value = f"{playlist.title} ({len(videos)})"
            page.update()

            data = []
            for video in videos:
                dct = OrderedDict()
                dct["URL"] = video.watch_url
                dct["Title"] = video.title
                dct["Channel"] = video.author
                data.append(dct)
                progress_bar.value += 1 / len(videos)
                page.update()

            filepath = Path(f"{playlist.title}.xlsx")
            XLSXFile(filepath).write(data)
            text_label.value = str(filepath.absolute())
            btn_export.disabled = False
            page.update()

        except Exception:
            btn_export.disabled = False
            progress_bar.value = 0
            text_label.value = "Произошла ошибка"
            page.update()
            raise

    page.title = "Export YouTube Playlist"
    page.vertical_alignment = ft.MainAxisAlignment.START
    page.window_width = 450
    page.window_height = 250
    page.window_resizable = False
    page.window_maximizable = False
    page.dark_theme = ft.Theme(color_scheme_seed=ft.colors.INDIGO)
    page.theme_mode = ft.ThemeMode.DARK

    text_field = ft.TextField(label="Link")
    btn_export = ft.ElevatedButton(text="Export", width=450, on_click=on_btn_export_click)
    progress_bar = ft.ProgressBar(value=0, color=ft.colors.RED)
    text_label = ft.Text(value="")

    page.add(text_field, btn_export, progress_bar, text_label)


if __name__ == '__main__':
    ft.app(target=main)
