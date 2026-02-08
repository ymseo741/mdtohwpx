from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
import os
import uuid
import md2hwpx
import shutil

app = FastAPI(title="MD to HWPX Converter API")

# Enable CORS for frontend integration
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

TEMP_DIR = "temp_conversions"
STATIC_DIR = "static"
os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(STATIC_DIR, exist_ok=True)

# Mount static files
app.mount("/assets", StaticFiles(directory=os.path.join(STATIC_DIR, "assets")), name="static")

@app.get("/", response_class=HTMLResponse)
async def serve_index():
    index_path = os.path.join(STATIC_DIR, "index.html")
    if os.path.exists(index_path):
        with open(index_path, "r", encoding="utf-8") as f:
            return f.read()
    return "Frontend not built yet. Run 'npm run build' in frontend directory."

@app.post("/convert")
@app.post("/api/convert") # Support both paths
async def convert_md(
    text: str = Form(None),
    file: UploadFile = File(None)
):
    if not text and not file:
        raise HTTPException(status_code=400, detail="No markdown content provided")

    content = ""
    if text:
        content = text
    else:
        content = (await file.read()).decode("utf-8")

    session_id = str(uuid.uuid4())
    output_filename = f"result_{session_id}.hwpx"
    output_path = os.path.join(TEMP_DIR, output_filename)

    try:
        # Perform conversion
        md2hwpx.convert_string(content, output_path)
        
        if not os.path.exists(output_path):
            raise HTTPException(status_code=500, detail="Conversion failed to generate file")

        return FileResponse(
            output_path, 
            filename="proposal.hwpx",
            media_type="application/octet-stream"
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/health")
def health_check():
    return {"status": "ok"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
