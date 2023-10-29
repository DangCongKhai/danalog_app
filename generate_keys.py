from pathlib import Path
import pickle
import streamlit_authenticator as stauth

names = ['Danalog']
user_name = ['DNL']
passwords = ['12345']


hashed_password = stauth.Hasher(passwords).generate()
file_path = Path(__file__).parent/"hashed_wd.pkl"

with file_path.open(mode='wb') as file:
    pickle.dump(hashed_password,file)