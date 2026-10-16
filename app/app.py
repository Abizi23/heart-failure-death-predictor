import pandas as pd
import streamlit as st
import numpy as np
import pickle
import base64



def get_binary_file_downloader_html(df):
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()

    href = f'<a href="data:file/csv;base64,{b64}" download="predictions.csv">Download Predictions CSV</a>'

    return href


st.title("DEATH_EVENT Predictor")
tab1,tab2,tab3=st.tabs(['Predict','Bulk Predict','Model Information'])

with tab1:

    age = st.number_input("Age", min_value=1, max_value=120, value=50)

    sex_option = st.selectbox("Sex", ["Male", "Female"])
    sex = 1 if sex_option == "Male" else 0

    smoking_option = st.selectbox("Smoking", ["No", "Yes"])
    smoking = 1 if smoking_option == "Yes" else 0


    anaemia_option = st.selectbox("Anaemia", ["No", "Yes"])
    anaemia = 1 if anaemia_option == "Yes" else 0

    diabetes_option = st.selectbox("Diabetes", ["No", "Yes"])
    diabetes = 1 if diabetes_option == "Yes" else 0

    hbp_option = st.selectbox("High Blood Pressure", ["No", "Yes"])
    high_blood_pressure = 1 if hbp_option == "Yes" else 0


    creatinine_phosphokinase = st.number_input(
        "Creatinine Phosphokinase (mcg/L)",
        min_value=0, value=100
    )

    ejection_fraction = st.number_input(
        "Ejection Fraction (%)",
        min_value=0, max_value=100, value=40
    )

    platelets = st.number_input(
        "Platelets (kiloplatelets/mL)",
        min_value=0.0, value=250000.0
    )

    serum_creatinine = st.number_input(
        "Serum Creatinine (mg/dL)",
        min_value=0.0, value=1.0
    )

    serum_sodium = st.number_input(
        "Serum Sodium (mEq/L)",
        min_value=0, value=137
    )

    time = st.number_input(
        "Follow-up period (days)",
        min_value=0, value=100
    )


    
    input_data = pd.DataFrame([{
        "age": age,
        "anaemia": anaemia,
        "creatinine_phosphokinase": creatinine_phosphokinase,
        "diabetes": diabetes,
        "ejection_fraction": ejection_fraction,
        "high_blood_pressure": high_blood_pressure,
        "platelets": platelets,
        "serum_creatinine": serum_creatinine,
        "serum_sodium": serum_sodium,
        "sex": sex,
        "smoking": smoking,
        "time": time
    }])


    algonames=['Logistic Regression','SVM']
    modelnames=['logistic_model.pkl','svm_model.pkl']

    
def predict_death_event(data):
    predictions = []

    for modelname in modelnames:
        with open(modelname, "rb") as file:
            model = pickle.load(file)

        prediction = model.predict(data)
        predictions.append(prediction[0])

    return predictions


if st.button("Submit"):
    result = predict_death_event(input_data)

    st.subheader("Results")

    for name, prediction in zip(algonames, result):
        st.write(f"{name}: {prediction}")


with tab2:
    st.title("Upload CSV File")

    st.subheader("Instructions to note before uploading the file:")

    st.info("""
    1. No missing (NaN) values are allowed.

    2. The CSV file must contain exactly these 12 input features
       with the correct column names:

       'age', 'anaemia', 'creatinine_phosphokinase', 'diabetes',
       'ejection_fraction', 'high_blood_pressure', 'platelets',
       'serum_creatinine', 'serum_sodium', 'sex', 'smoking', 'time'.

    3. Check the spelling of the feature names. Column names must
       match exactly.

    4. Feature value conventions:

       - age: age of the patient [years]
       - anaemia: 1 = Yes, 0 = No
       - creatinine_phosphokinase: creatinine phosphokinase level [mcg/L]
       - diabetes: 1 = Yes, 0 = No
       - ejection_fraction: percentage of blood leaving the heart
         at each contraction [%]
       - high_blood_pressure: 1 = Yes, 0 = No
       - platelets: platelet count [kiloplatelets/mL, as represented
         in the dataset]
       - serum_creatinine: serum creatinine level [mg/dL]
       - serum_sodium: serum sodium level [mEq/L]
       - sex: 1 = Male, 0 = Female
       - smoking: 1 = Yes, 0 = No
       - time: follow-up period [days]

    5. Do not include the target column 'DEATH_EVENT' among the
       input features.

    6. The CSV file should contain one patient per row. You can
       upload multiple rows for bulk prediction.

    7. Save the file in CSV format (.csv) before uploading.
    """)

    uploaded_file=st.file_uploader("Upload a CSV file",type=['csv'])

    if uploaded_file is not None:
        input_data=pd.read_csv(uploaded_file)
        model=pickle.load(open('LogisticRegression.pkl','rb'))

        expected_columns = [
    "age",
    "anaemia",
    "creatinine_phosphokinase",
    "diabetes",
    "ejection_fraction",
    "high_blood_pressure",
    "platelets",
    "serum_creatinine",
    "serum_sodium",
    "sex",
    "smoking",
    "time"
]       
        if set(expected_columns).issubset(input_data.colums):
             

           input_data['Prediction LR']=''

           for i in range(len(input_data)):
              arr=input_data.iloc[i,:-1].values
              input_data['Prediction'][i]=model.predict([arr])[0]
           input_data.to_csv('PredictedDeathEvent.csv')


           st.subheader('Predictions:')
           st.write(input_data)

           st.markdown (get_binary_file_downloader_html(input_data),unsafe_allow_html=True)
        else:
            st.warning('Please make sure the uploaded CSV file has the correct columns.')
           

    else:
        st.info('Upload a CSV file to get predicitons.')




with tab3:
    import plotly.express as px
    data={'Logistic Regression':80.97, "SVM":90}
    models= list(data.keys())
    Accuracies=list(data.values())
    df=pd.DataFrame(list(zip(models,Accuracies)),columns=['Models','Accuracies'])
    fig=px.bar(df,y='Accuracies',x='Models')
    st.plotly_chart(fig)