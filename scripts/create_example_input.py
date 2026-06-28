"""Create an example input Excel file for testing."""

from openpyxl import Workbook


def create_example_input(path: str = "example_input.xlsx") -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Quotations"

    ws.append(["Document Number"])
    ws.append(["20001234"])
    ws.append(["20005678"])
    ws.append(["20009012"])

    wb.save(path)
    print(f"Created example input file: {path}")


if __name__ == "__main__":
    create_example_input()
