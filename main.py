from typing import List
from fastapi import FastAPI,File, UploadFile, HTTPException, Body
from pydantic import BaseModel, EmailStr
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
import mysql.connector
from starlette.responses import FileResponse
from fastapi.responses import StreamingResponse
import pandas as pd
import io
from docx import Document

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Update this with your frontend URL in production
    allow_credentials=True,
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
)
# Define Pydantic models for request and response data
class User(BaseModel):
    email: EmailStr
    password: str

class UserInDB(User):
    id: int


conn = mysql.connector.connect(
    host="localhost",
    user="root",
    password="Password12",
    database="excelstore"
)

cursor = conn.cursor()


# Function to register a new user
@app.post('/register')
async def register(user: User):
    try:
       
        # Execute the query to insert a new user into the users table
        cursor = conn.cursor()
        cursor.execute("INSERT INTO users (email, password) VALUES (%s, %s)", (user.email, user.password))
        conn.commit()

        # Close the cursor and connection
        cursor.close()
        

        return {'message': 'User registered successfully'}

    except mysql.connector.Error as error:
        # Raise HTTPException if an error occurs
        raise HTTPException(status_code=400, detail="Error registering user")

# Function to authenticate a user during login
@app.post('/login')
async def login(user: User):
    try:
        
        # Execute the query to check if the user exists
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM users WHERE email = %s AND password = %s", (user.email, user.password))
        user_row = cursor.fetchone()

        # Close the cursor and connection
        cursor.close()
        

        if user_row:
            return {'message': 'Login successful'}
        else:
            raise HTTPException(status_code=401, detail="Invalid email or password")

    except mysql.connector.Error as error:
        # Raise HTTPException if an error occurs
        raise HTTPException(status_code=400, detail="Error logging in")
@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    try:
        # Check if file was provided
        if not file:
            raise HTTPException(status_code=400, detail="No file provided")

        # Read the Excel file
        file_contents = await file.read()  # Read file contents as bytes
        excel_data = pd.read_excel(io.BytesIO(file_contents))  # Convert to BytesIO and read with Pandas
        
        # Convert Excel data to binary
        excel_binary_data = excel_data.to_csv(None, index=False, header=False).encode()
        
        file_name = file.filename

        # Insert binary data into MySQL table
        insert_query = "INSERT INTO excel_files (file_name, file_data) VALUES (%s,%s)"
        cursor.execute(insert_query, (file_name, excel_binary_data))
        conn.commit()  # Commit changes to the database

        return JSONResponse(content={"message": "File uploaded successfully"}, status_code=200)
    except Exception as e:
        # Log the exception for debugging
        print(f"Error during file processing: {e}")
        raise HTTPException(status_code=500, detail="Internal Server Error")

@app.get('/files/')
async def get_stored_files():
    try:
        # Execute query to retrieve file names from the database
        cursor.execute("SELECT file_name FROM excel_files")
        files = [row[0] for row in cursor.fetchall()]  # Extract file names from result
        return {"files": files}
    except Exception as e:
        print(f"Error fetching stored files: {e}")
        raise HTTPException(status_code=500, detail="Internal Server Error")
    


@app.get('/download/{file_name}')
async def download_file(file_name: str):
    try:
        # Execute query to retrieve file data from the database
        cursor.execute("SELECT file_data FROM excel_files WHERE file_name = %s", (file_name,))
        file_data = cursor.fetchone()

        if file_data:
            # Extract file content
            excel_binary_data = file_data[0]

            # Create a StreamingResponse to stream the file
            return StreamingResponse(io.BytesIO(excel_binary_data), media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', headers={"Content-Disposition": f'attachment; filename="{file_name}"'})
        else:
            raise HTTPException(status_code=404, detail="File not found")

    except Exception as e:
        print(f"Error downloading file: {e}")
        raise HTTPException(status_code=500, detail="Internal Server Error")



@app.post('/concatenate-generate-download')
async def concatenate_generate_download(selected_files: List[str] = Body(...)):
    try:
        # Initialize a new Word document
        doc = Document()

        # Loop through each selected file
        for file_name in selected_files:
            cursor.execute("SELECT file_data FROM excel_files WHERE file_name = %s", (file_name,))
            file_data = cursor.fetchone()

            if file_data:
                # Extract file content
                excel_binary_data = file_data[0]

                # Convert Excel data to DataFrame
                df = pd.read_csv(io.BytesIO(excel_binary_data))

                # Add DataFrame content to the Word document
                table = doc.add_table(df.shape[0] + 1, df.shape[1])
                for i in range(df.shape[0] + 1):
                    for j in range(df.shape[1]):
                        cell = table.cell(i, j)
                        if i == 0:
                            cell.text = df.columns[j]
                        else:
                            cell.text = str(df.iloc[i - 1, j])

                # Consume the result set
                cursor.fetchall()

        # Save the Word document
        file_path = 'concatenated_file.docx'
        doc.save(file_path)

        # Send the generated Word file to the client for download
        return FileResponse(file_path, filename=file_path)

    except Exception as e:
        print(f"Error generating concatenated Word file: {e}")
        raise HTTPException(status_code=500, detail="Internal Server Error")

    



# Close the MySQL connection (outside of the endpoint)
@app.on_event("shutdown")
def shutdown_event():
    conn.close()

if __name__ == '__main__':
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)
