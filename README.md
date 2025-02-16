# Gift Voucher Maker

## How to use
1. Install uv first by opening PowerShell, then running:
    ```
    winget install --id=astral-sh.uv  -e
    ```
1. Ensure that the `Hari Raya Vouchers.csv` file is furnished.
1. The `template_hariraya.pptx` can be modified to taste, as long as there are text boxes containing "Voucher Code", "Voucher Link" and "QR Code" (If not found, there will be an error).
1. Within the project folder run `uv run python main.py`, there will be an `output.pptx` file with desired output.