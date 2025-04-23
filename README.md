# Product Classification Tool

This is a tool that allows you to classify product titles based on predefined rules and logic. It is built using Python, Streamlit, and Openpyxl, and it can read and process Excel files for product classification.

## Features

- **Classify Product Titles:** Classify product titles based on rules defined in an Excel sheet.
- **Keyword Matching:** Supports both AND and OR logic for matching keywords in product titles.
- **Highlight Matching Keywords:** The tool highlights the matching keywords in the product titles.
- **Result Export:** The results are saved in an Excel file, which can be downloaded.

## Requirements

To run this project locally, you need to install the required dependencies. You can install them using the `requirements.txt` file:

```bash
pip install -r requirements.txt
```

The `requirements.txt` includes the following libraries:
- `pandas` – for data manipulation.
- `openpyxl` – for reading and writing Excel files.
- `streamlit` – for building and running the web app.

## Usage

### Local Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/your_username/product-classification-tool.git
   cd product-classification-tool
   ```

2. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Run the Streamlit app:
   ```bash
   streamlit run app.py
   ```

4. Open your browser and navigate to `http://localhost:8501` to use the app.

### Using the Streamlit App

1. Upload your Excel file that contains the product data and classification rules. The file should have two sheets:
   - **Product detail:** Should contain a column named `TITLE` with product titles to be classified.
   - **Rules:** Should contain the classification rules with three columns: `Rule`, `Include`, and `Exclude`.
   
2. The app will classify the product titles based on the provided rules:
   - **Include:** Defines the keywords that must appear (with AND/OR logic).
   - **Exclude:** Defines the keywords that should be excluded.
   
3. Once the classification is done, the results will be displayed, and you can download the classified product list as an Excel file.

### Streamlit Deployment

If you want to deploy this app online, you can use Streamlit's cloud platform. Follow these steps:

1. Push your code to GitHub (if not already done).
2. Visit [Streamlit Cloud](https://streamlit.io/cloud) and sign in.
3. Create a new app, linking your GitHub repository.
4. The app will be deployed automatically, and you will get a link to share with others.

## Files

- `app.py`: The main Streamlit application file.
- `requirements.txt`: Lists the Python dependencies needed to run the app.
- `README.md`: This file with project details and instructions.

## License

This project is licensed under the MIT License – see the [LICENSE](LICENSE) file for details.

