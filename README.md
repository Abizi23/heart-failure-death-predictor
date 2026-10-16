# Heart Failure Death Event Prediction

## Overview
This project uses machine learning to predict the risk of a death event among patients with heart failure based on clinical features. It compares two classification algorithms: **Logistic Regression** and **Support Vector Machine (SVM)**.

## Features
- Data exploration and model development using Google Colab.
- Prediction using clinical variables such as age, anaemia, diabetes, ejection fraction, serum creatinine, and serum sodium.
- Interactive web application developed with Streamlit.
- Model accuracy comparison using Plotly visualizations.
- Individual patient predictions and support for bulk CSV prediction.

## Technologies Used
- Python
- Pandas and NumPy
- Scikit-learn
- Plotly Express
- Streamlit
- Google Colab

## Project Structure
- `notebooks/` — Model development and evaluation.
- `app/` — Streamlit application and trained machine learning models.
- `requirements.txt` — Python dependencies.

## How to Run
1. Clone this repository.
2. Install the required packages:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the Streamlit application from the project root:
   ```bash
   streamlit run app/app.py
   ```

## Disclaimer
This project is intended for educational and research purposes only. Predictions are not medical diagnoses and should not be used as a substitute for professional medical advice.

