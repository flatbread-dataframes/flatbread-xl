from pathlib import Path

import pandas as pd

from uvmlib.config import PATHS, DataBron

class DocSheetBuilder:
    """Maak een documentatiesheet in een Excel workbook."""

    def __init__(
        self,
        titel: str,
        bronnen: list[DataBron],
        script: Path,
        doel: str,
        methode: str,
        sheets_info: dict[str, str],
        sheet_name: str = "documentatie",
        tab_kleur: str | None = "#4472C4",
        timestamp: str|None = None,
    ) -> None:
        self.titel = titel
        self.timestamp = timestamp or f"{pd.Timestamp.now():%d-%m-%Y %H:%M}"
        self.bronnen = bronnen
        self.script = script
        self.doel = doel
        self.methode = methode
        self.sheets_info = sheets_info
        self.sheet_name = sheet_name
        self.tab_kleur = tab_kleur

    def build(self, writer: pd.ExcelWriter) -> None:
        workbook = writer.book
        worksheet = workbook.add_worksheet(self.sheet_name)

        worksheet.hide_gridlines(2)
        worksheet.set_tab_color(self.tab_kleur)

        # Stel kolombreedtes in
        COL_WIDTHS = [
            20,     # Label kolom
            60,     # Content kolom
            100,    # Method textbox kolom
        ]
        for i, width in enumerate(COL_WIDTHS):
            worksheet.set_column(i, i, width)

        # Voeg cell-formats toe
        wrap_format = workbook.add_format({'text_wrap': True})
        self.label_format = workbook.add_format({
            'bold': True,
            'valign': 'top'
        })
        self.title_format = workbook.add_format({
            'bold': True,
            'font_size': 14,
            # 'align': 'center'
        })

        current_row = 0

        # Titel
        worksheet.merge_range(
            first_row   = current_row,
            first_col   = 0,
            last_row    = current_row,
            last_col    = 2,
            data        = self.titel,
            cell_format = self.title_format
        )
        current_row += 2

        # Methode
        worksheet.write(2, 2, "Methode", self.label_format)
        worksheet.insert_textbox(
            3, 2,
            self.methode,
            {
                'width': COL_WIDTHS[2] * 7,
                'height': 720,
                'font': {
                    'name': 'Consolas',
                    'size': 10
                },
                'text_wrap': True,
                'border': {'none': True},
                'align': {'vertical': 'top'},
            }
        )

        # Overige onderdelen
        def add_section(label: str, content: str):
            worksheet.write(current_row, 0, label, self.label_format)
            worksheet.write(current_row, 1, content, wrap_format)
            return current_row + 2

        current_row = add_section("Datum aangemaakt", self.timestamp)

        # Bronnen
        worksheet.write(current_row, 0, "Bronbestanden", self.label_format)
        current_row += 1
        for bron in self.bronnen:
            worksheet.write(current_row, 0, bron.key.partition('.')[-1])
            worksheet.write(current_row, 1, str(get_path_binnen_projectstructuur(bron.filepath)))
            current_row += 1
        current_row += 1

        # Script
        script_path = get_path_binnen_projectstructuur(self.script, PATHS.lib)
        current_row = add_section("Script", str(script_path))

        # Doel
        current_row = add_section("Doel", self.doel)

        # Sheets info
        worksheet.write(current_row, 0, "Sheets", self.label_format)
        current_row += 1

        sheet_name_format = workbook.add_format({'valign': 'top'})
        for sheet_name, description in self.sheets_info.items():
            worksheet.write(current_row, 0, sheet_name, sheet_name_format)
            worksheet.write(current_row, 1, description, wrap_format)
            current_row += 1


def ligt_binnen_parent(child: Path, parent: Path) -> bool:
    """
    Controleer of de child folder of bestand binnen de parent folder ligt.

    Parameters
    ----------
    child : Path
        Te controleren folder of bestand
    parent : Path
        Folder waarin child aanwezig zou moeten zijn

    Returns
    -------
    bool
        True als child in parent aanwezig is, anders False

    Notes
    -----
    Houdt rekening met de inconsequente manier waarop Windows met netwerkpaden omgaat:
    - Gekoppelde netwerkschijf (bijv. 'O:\\some\\path')
    - UNC-pad (bijv. '\\\\server\\share\\some\\path')

    Door deze inconsistentie kan niet op de standaard manier bepaald worden of een bestand of folder binnen een andere folder ligt.
    """
    child_parts = child.resolve().parts
    parent_parts = parent.resolve().parts
    return len(parent_parts) <= len(child_parts) and child_parts[:len(parent_parts)] == parent_parts


def get_path_binnen_projectstructuur(
    full_path: Path,
    project_path: Path = PATHS.project,
) -> Path:
    """
    Zet een absoluut pad om naar de relatieve locatie van dat pad binnen de projectstructuur.

    Parameters
    ----------
    full_path : Path
        Om te zetten pad
    project_path : Path, optional
        Project root pad, standaard PATHS.project

    Returns
    -------
    Path
        Relatief pad binnen projectstructuur

    Raises
    ------
    ValueError
        Als full_path niet binnen de project_path ligt

    Notes
    -----
    Met netwerkschijf:
    >>> get_path_binnen_projectstructuur(Path('O:/project/data/file.csv'))
    Path('data/file.csv')

    Met UNC-pad:
    >>> get_path_binnen_projectstructuur(Path('\\\\server\\share\\project\\data\\file.csv'))
    Path('data/file.csv')
    """
    full = full_path.resolve()
    base = project_path.resolve()

    if not ligt_binnen_parent(full, base):
        raise ValueError(f"{full_path} ligt niet binnen {project_path}")

    return full.relative_to(base)
