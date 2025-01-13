from fastapi import FastAPI
from excel_routes import excel_router

# Initialize FastAPI app
app = FastAPI()
app.include_router(excel_router)

# Run the application using uvicorn
if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="127.0.0.1", port=8000, reload=True)
