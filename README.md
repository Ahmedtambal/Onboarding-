# My Streamlit App

This project is a Streamlit application designed for streamlined onboarding of new employees. It allows users to upload employee details files (in DOCX or PDF format) and an Excel master record. The application processes the uploaded files, extracts relevant data, and updates the master record accordingly.

## Project Structure

```
my-streamlit-app
├── app.py          # Main Streamlit application code
├── logic.py        # Logic for parsing documents and handling data
├── requirements.txt # Python dependencies
└── README.md       # Project documentation
```

## Installation

To set up the project, follow these steps:

1. Clone the repository:

   ```
   git clone <repository-url>
   cd my-streamlit-app
   ```

2. Create a virtual environment (optional but recommended):

   ```
   python -m venv venv
   source venv/bin/activate  # On Windows use `venv\Scripts\activate`
   ```

3. Install the required dependencies:

   ```
   pip install -r requirements.txt
   ```

## Usage

To run the Streamlit application, use the following command:

```
streamlit run app.py
```

Once the application is running, you can access it in your web browser at `http://localhost:8501`.

## Features

- Upload employee details files in DOCX or PDF format.
- Upload an Excel master record.
- View extracted employee data and the current master record.
- Download the updated master file after processing.

## Deployment

To deploy the application on Streamlit Sharing, follow these steps:

1. Push your code to a GitHub repository.
2. Go to [Streamlit Sharing](https://streamlit.io/sharing) and sign in.
3. Click on "New app" and select your GitHub repository.
4. Choose the branch and the file path to `app.py`.
5. Click "Deploy" to launch your application.

## License

This project is licensed under the MIT License. See the LICENSE file for details.