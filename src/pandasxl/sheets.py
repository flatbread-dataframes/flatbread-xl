from pathlib import Path

import pandas as pd

from uvmlib.export.docsheet import DocSheetBuilder


class Sheet:
    """
    Sheet in een Excel workbook met metadata.

    Parameters
    ----------
    naam : str
        Naam van de sheet
    data : pd.DataFrame | pd.Series
        Data die in de sheet moet worden weggeschreven
    omschrijving : str
        Omschrijving van de inhoud van de sheet
    toon_index : bool
        Optie om de index wel/niet op te nemen in de output, default True
    tab_kleur : str | None, optional
        Hex kleurcode voor de sheet (bijv. '#70AD47'), default None
    """
    def __init__(
        self,
        naam: str,
        data: pd.DataFrame|pd.Series,
        omschrijving: str,
        toon_index: bool = True,
        tab_kleur: str|None = None,
    ) -> None:
        self.naam = naam
        self.data = data if isinstance(data, pd.DataFrame) else data.to_frame()
        self.omschrijving = omschrijving
        self.toon_index = toon_index
        self.tab_kleur = tab_kleur


class OutputSheet(Sheet):
    """Sheet met OUTPUT data."""
    def __init__(
        self,
        data: pd.DataFrame|pd.Series,
        omschrijving: str,
        toon_index: bool = True,
        tab_kleur = "#70AD47",  # Medium green
    ) -> None:
        super().__init__(
            naam = "OUTPUT",
            data = data,
            omschrijving = omschrijving,
            toon_index = toon_index,
            tab_kleur = tab_kleur,
        )


class InterimSheet(Sheet):
    """Sheet met INTERIM resultaten."""
    def __init__(
        self,
        data: pd.DataFrame|pd.Series,
        omschrijving: str,
        toon_index: bool = True,
        tab_kleur = "#FFC000",  # Gold
    ) -> None:
        super().__init__(
            naam = "INTERIM",
            data = data,
            omschrijving = omschrijving,
            toon_index = toon_index,
            tab_kleur = tab_kleur,
        )


class DocSheet(Sheet):
    """Sheet met documentatie."""
    def __init__(
        self,
        titel: str,
        bronnen: list[Path],
        script: Path,
        doel: str,
        methode: str,
        sheets: list[Sheet],
        tab_kleur = "#4472C4",  # Professional blue
        timestamp: str|None = None,
    ) -> None:
        super().__init__(
            naam = "documentatie",
            data = pd.DataFrame(),
            omschrijving = "Documentatie",
            tab_kleur = tab_kleur,
        )

        self.builder = DocSheetBuilder(
            titel = titel,
            bronnen = bronnen,
            script = script,
            doel = doel,
            methode = methode,
            sheets_info = {sheet.naam:sheet.omschrijving for sheet in sheets},
            sheet_name = self.naam,
            tab_kleur = self.tab_kleur,
            timestamp = timestamp,
        )

    def write_to_excel(self, writer: pd.ExcelWriter) -> None:
        self.builder.build(writer)


class MultiTableSheet(Sheet):
    """
    Sheet die meerdere dataframes met titel kan wegschrijven naar een excelsheet.

    Parameters
    ----------
    naam : str
        Naam van de sheet
    tables : list[tuple[str, pd.DataFrame | pd.Series, str | None]]
        Lijst met tabellen (titel, df, caption) die in de sheet moeten worden weggeschreven.
        Caption is optioneel en kan None zijn.
    omschrijving : str
        Omschrijving van de inhoud van de sheet
    spacing : int, optional
        Aantal lege rijen tussen tabellen, default 2
    toon_index : bool
        Optie om de index wel/niet op te nemen in de output, default True
    tab_kleur : str | None, optional
        Hex kleurcode voor de sheet (bijv. '#70AD47'), default None
    """
    def __init__(
        self,
        naam: str,
        tables: list[tuple[str, pd.DataFrame | pd.Series, str | None]],
        omschrijving: str = '',
        spacing: int = 1,
        toon_index: bool = True,
        tab_kleur: str | None = None,
    ) -> None:
        self.tables = []
        for item in tables:
            # Prepareer optionele caption
            if len(item) == 2:
                title, df = item
                caption = None
            else:
                title, df, caption = item

            if isinstance(df, pd.Series):
                df = df.to_frame()

            self.tables.append((title, df, caption))

        self.spacing = spacing
        super().__init__(
            naam = naam,
            data = pd.DataFrame(),
            omschrijving = omschrijving,
            toon_index = toon_index,
            tab_kleur = tab_kleur,
        )

    def write_to_excel(self, writer: pd.ExcelWriter) -> None:
        """Write all tables to the worksheet with titles."""
        sheet = writer.book.add_worksheet(self.naam)
        title_format = writer.book.add_format({'bold': True})
        caption_format = writer.book.add_format({'italic': True})

        if self.tab_kleur:
            sheet.set_tab_color(self.tab_kleur)

        current_row = 0
        for title, df, caption in self.tables:
            # Bereken breedte voor titel/caption
            index_width = df.index.nlevels if self.toon_index else 0
            total_width = index_width + len(df.columns)

            # Schrijf titel weg
            sheet.merge_range(
                first_row=current_row,
                first_col=0,
                last_row=current_row,
                last_col=total_width - 1,
                data=title,
                cell_format=title_format
            )

            # Schrijf data weg
            df.to_excel(
                writer,
                sheet_name=self.naam,
                startrow=current_row + 1,
                startcol=0,
                index=self.toon_index
            )

            last_row = current_row + 1 + len(df) + df.columns.nlevels

            # Add caption if provided
            if caption:
                sheet.merge_range(
                    first_row=last_row,
                    first_col=0,
                    last_row=last_row,
                    last_col=total_width - 1,
                    data=caption,
                    cell_format=caption_format
                )
                # Adjust for caption
                last_row += 1

            # Move to next position
            current_row = last_row + self.spacing

        sheet.autofit()
