# If you see "Oh no. Error running app" on Streamlit Cloud

1. **Check Logs (most important)**  
   In [share.streamlit.io](https://share.streamlit.io) → your app → **Manage app** → **Logs**.  
   The full Python traceback appears there and shows the exact error.

2. **Main file path**  
   In app settings, **Main file path** must be exactly `app.py` (or `app_simple.py` to test deployment).

3. **Python version**  
   Use **Python 3.12** for this app. Python 3.13 can cause "Oh no" / dependency errors. In App settings → General → Python version, select **3.12** and Save.

4. **Test with minimal app**  
   Set Main file path to `app_simple.py`, save, and reopen the app.  
   - If you see "TDR Portal – deployment OK", deployment is fine; switch back to `app.py`.  
   - If you still see "Oh no", the issue is repo/branch or main file path.

5. **Repo and branch**  
   Confirm the app is using the correct GitHub repo and branch (e.g. `main`).
