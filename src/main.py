from datetime import timedelta
import flet as ft
from xls_to_ass import convert_datetime, convert_worksheet_to_ass, load_excel_file

STYLES = {
    "title": ft.TextStyle(size=30, weight=ft.FontWeight.BOLD),
    "subtitle": ft.TextStyle(size=16, weight=ft.FontWeight.BOLD),
}

DT_MAX_DISPLAY_ROWS = 4
DT_PLACEHOLDER_COLS = 5
def generate_placeholder_datatable() -> ft.DataTable:
    rows = [ft.DataRow([ft.DataCell(ft.Text("    ")) for i in range(DT_PLACEHOLDER_COLS)]) for j in range(DT_MAX_DISPLAY_ROWS)]
    cols = [ft.DataColumn(ft.Text("     ")) for i in range(DT_PLACEHOLDER_COLS)]
    return ft.DataTable(columns=cols, rows=rows, data_row_min_height=0, expand=True)

def format_ws_to_datatable(ws: list[list[str]], has_headers: bool = True) -> tuple[ft.DataTable, list[tuple[int, str]]]:
    i = 0
    rows = []
    cols = []
    dd_options = []

    for r in ws:
        if not cols:
            if has_headers:
                col_names = r
                for j in range(len(col_names)):
                    if not col_names[j]:
                        col_names[j] = f"Column {j+1}"                
                cols = [ft.DataColumn(ft.Text(s)) for s in col_names]
                dd_options = [(str(k), s) for k, s in zip(range(len(r)), col_names)]
                continue
            else:
                col_names = [f"Column {k+1}" for k in range(len(r))]
                cols = [ft.DataColumn(ft.Text(s)) for s in col_names]
                dd_options = [(str(k), s) for k, s in zip(range(len(r)), col_names)]
        if i == DT_MAX_DISPLAY_ROWS:
            break
        rows.append(ft.DataRow([ft.DataCell(
            ft.Text(c, size=12)
            ) for c in r]))
        i = i + 1
    
    dd_options.insert(0, ("-1", " "))

    data_table = ft.DataTable(columns=cols, rows=rows, data_row_min_height=0, expand=True)

    return data_table, dd_options

def configure_timestamp_render_page(page: ft.Page) -> None:
    c = ft.Column([
        ft.Text("Convert XLS to ASS", size=30, weight=ft.FontWeight.BOLD),
        ft.Text("Step 4: Configure timestamp properties", size=20),
        ft.Divider(),
        ft.Row([
            ft.Text
        ])
    ])
    pass


def create_column_dd(label: str, hint) -> ft.Dropdown:
    return ft.Dropdown(label=label, expand=1, value=-1, hint_text=hint)

def load_xls_dialog_on_result(e: ft.FilePickerResultEvent, load_xls_btn: ft.Button, dd: ft.Dropdown, load_dd_btn: ft.Button):
    page = e.page

    if not e.files:
        dd.value = ""
        dd.disabled = True
        load_dd_btn.disabled = True
        page.update()
        return

    fn = e.files[0].name
    try:
        wb = load_excel_file(fn)
    except Exception as e:
        print(e)
        dd.value = ""
        dd.disabled = True
        load_dd_btn.disabled = True
        page.update()
        return

    load_xls_btn.text = fn
    dd.options = [ft.DropdownOption(s, s) for s in list(wb.keys())]
    dd.value = list(wb.keys())[0]
    dd.disabled = False
    load_dd_btn.disabled = False
    page.data = wb
    page.update()

def load_worksheet_on_click(e, ws_name: str, has_headers: bool, dt: ft.Container, dd_list: list[ft.Dropdown]):
    page = e.page
    ws = page.data[ws_name]
    data_table, dd_options = format_ws_to_datatable(ws, has_headers)
    dt.content = data_table
    dt.update()
    for dd in dd_list:
        dd.options = [ft.DropdownOption(key=k, text=s) for k, s in dd_options]
        dd.update()
    page.update()

def parse_col_value(s: str):
    i = int(s)
    return i if i >= 0 else None

def save_ass_dialog_on_result(e, ws_dd: ft.Dropdown, start_dd: ft.Dropdown, end_dd: ft.Dropdown, dialogue_dd: ft.Dropdown,
                              actor_dd: ft.Dropdown, track_dd: ft.Dropdown, italics_dd: ft.Dropdown, headers_checkbox: ft.Checkbox,
                              is_timecode_dd: ft.Dropdown, framerate: ft.TextField, shift: ft.TextField, scale: ft.TextField):
    if not e.path:
        return
    
    output_fn = e.path
    if output_fn[-4:] != ".ass":
        output_fn = output_fn + ".ass"
    page = e.page
    ws = page.data[ws_dd.value]

    has_headers = headers_checkbox.value
    start_col = parse_col_value(start_dd.value)
    end_col = parse_col_value(end_dd.value)
    dialogue_col = parse_col_value(dialogue_dd.value)
    actor_col = parse_col_value(actor_dd.value)
    track_col = parse_col_value(track_dd.value)
    italics_col = parse_col_value(italics_dd.value)

    if start_col is None and dialogue_col is None:
        print("error")
        return

    args = {}
    if is_timecode_dd.value is None:
        print("error")
        return
    args["is_timecode"] = is_timecode_dd.value == "True"
    try:
        if framerate.value:
            args["framerate"] = float(framerate.value)
        if shift.value:
            args["shift"] = convert_datetime(shift.value)
        if scale.value:
            args["scale"] = float(scale.value)
    except Exception as e:
        print(e)
        return
    
    doc = convert_worksheet_to_ass(ws, start_col=start_col, end_col=end_col, dialogue_col=dialogue_col, actor_col=actor_col,
                             track_col=track_col, italics_col=italics_col, has_headers=has_headers, convert_timestamp_args=args)
    with open(output_fn, mode="w", encoding="utf_8_sig") as f:
        doc.dump_file(f)




def main(page: ft.Page):
    page.title = "Convert XLS to ASS"
    page.vertical_alignment = ft.MainAxisAlignment.START
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.window.width = 1000
    page.window.height = 800
    page.theme_mode = ft.ThemeMode.SYSTEM
    page.padding = 20
    page.scroll = ft.ScrollMode.ALWAYS

    load_xls_fp = ft.FilePicker()
    save_ass_fp = ft.FilePicker()

    load_xls_btn = ft.Button("Load file...", expand=1)
    worksheet_dd = ft.Dropdown(disabled=True)
    header_checkbox = ft.Checkbox("Has headers", value=True)
    load_worksheet_btn = ft.Button("Load worksheet...", expand=1)
    
    data_table = ft.Container(generate_placeholder_datatable(), expand=True)
    
    start_time_dd = create_column_dd("Start time", "Selects the column that represents the start time for the subtitle. Usually has data in a format like 01:02:03.04")
    end_time_dd =   create_column_dd("End time", "Selects the column that represents the start time for the subtitle. Usually has data in a format like 01:02:03.04")
    dialogue_dd =   create_column_dd("Dialogue", "Selects the column that represents the text to be shown on screen. Some files have multiple dialogue columns - choose one.")
    actor_dd =      create_column_dd("Actor", "Selects the column that represents the actor/character saying the line. Optional, although relatively common. This data is stored on the Actor field of each line on the ASS file.")
    track_dd =      create_column_dd("Track", "Selects the column that represents the 'track' of the subs. Optional. Many vendors separate dialogue and signs events in two different files, usually named tracks 'A' and 'B'. If this data is present on the XLS, the convertor creates separate styles for each track.")
    italics_dd =    create_column_dd("Italics", "Selects the column that represents the italicization of the subs. Optional. Some vendors add a column where they indicate with a * sign or something similar whenever the line must be shown in italics or not. Selecting this column adds a {\\i1} tag to the start of the line whenever the column is not empty.")
    dd_list: list[ft.Dropdown] = [start_time_dd, end_time_dd, dialogue_dd, actor_dd, track_dd, italics_dd]
    
    framerate_mode_dd = ft.Dropdown( label="Timestamp Type", value="True", options=[
        ft.DropdownOption("True", "Timecode"),
        ft.DropdownOption("False", "Seconds")
    ], hint_text="Used to determine whether the subs timestamps are in timecode format or not. In timecode format, the sub-second digits represent exact frames - otherwise, they represent fractions of a second. If you're unsure which option to pick, open your file and take a look at the first dozen of lines or so. If the last couple of digits are always under 23, 24 or 29, that probably means that the timestamp is in timecode format for 24, 25 or 30fps content.")
    framerate_textbox = ft.TextField(expand=1, label="Timecode Framerate", value="24", hint_text="Determines the framerate used to convert timecode-format timestamps. Unused if in Seconds mode. Please read the hint for Timestamp Format to know more.")
    shift_textbox = ft.TextField(expand=1, label="Timecode shift", value="", hint_text="Determines by how much each timecode should be shifted when converting. Can be positive or negative. Must have format hh:mm:ss.cc. It's very common to use -01:00:00.00 to remove the initial timecode value from the subs.")
    scale_textbox = ft.TextField(expand=1, label="Timecode scale", value="", hint_text="Determines by how much each timecode should be multiplied when converting. Must be a positive float number. Commonly used to adjust from 24fps to 23.976fps (and vice-versa), 30fps to 29.97fps, 25fps to 24fps, and so on. Scaling is applied after shifting.")
    save_as_btn = ft.Button("Export to ASS...")

    load_xls_btn.on_click = lambda _: load_xls_fp.pick_files("Load XLS file",
                                                                allowed_extensions=["xls", "xlsx", "xlsm"], 
                                                                allow_multiple=False)
    load_worksheet_btn.on_click = lambda e: load_worksheet_on_click(e, worksheet_dd.value, header_checkbox.value, data_table, dd_list)

    save_as_btn.on_click = lambda _: save_ass_fp.save_file("Save as...", allowed_extensions=["ass"])
    load_xls_fp.on_result = lambda e: load_xls_dialog_on_result(e, load_xls_btn, worksheet_dd, load_worksheet_btn)
    save_ass_fp.on_result = lambda e: save_ass_dialog_on_result(e, worksheet_dd, start_time_dd, end_time_dd, dialogue_dd, actor_dd, track_dd,
                                                          italics_dd, header_checkbox, framerate_mode_dd, framerate_textbox,
                                                          shift_textbox, scale_textbox)

    c = ft.Column([
        ft.Text("Convert XLS to ASS", style=STYLES["title"]),
        ft.Row([
            ft.Text("Step 1: Load the XLS", style=STYLES["subtitle"]),
            load_xls_btn
        ], expand=True),
        ft.Divider(),
        ft.Row([
            ft.Text("Step 2: Select the worksheet", style=STYLES["subtitle"]),
            worksheet_dd,
            header_checkbox,
            load_worksheet_btn,
        ], expand=True),
        ft.Divider(),
        ft.Row([
            data_table
        ], expand=True),
        ft.Divider(),
        ft.Row([
            ft.Text("Step 3: Configure column meaning", style=STYLES["subtitle"]),
            *dd_list
        ], expand=True),
        ft.Divider(),
        ft.Row([
            ft.Text("Step 4: Configure timestamp properties", style=STYLES["subtitle"]),
            framerate_mode_dd,
            framerate_textbox,
            shift_textbox,
            scale_textbox
        ], expand=True),
        ft.Divider(),
        ft.Row([ft.Text("Step 5: Export to ASS", style=STYLES["subtitle"]), save_as_btn])
    ])
    page.controls.append(c)
    page.overlay.append(load_xls_fp)
    page.overlay.append(save_ass_fp)
    page.update()

if __name__ == "__main__":
    ft.app(target=main)