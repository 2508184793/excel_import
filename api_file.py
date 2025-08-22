from typing import Union

from fastapi import FastAPI, UploadFile, File
import os

app = FastAPI()

# 确保文件存储目录存在
FILE_DIR = "FileDir"
if not os.path.exists(FILE_DIR):
    os.makedirs(FILE_DIR)


# 测试接口

@app.get("/items/{item_id}", summary='获取文件信息')
def read_item(item_id: int, q: Union[str, None] = None):
    return {"item_id": item_id, "q": q}


from fastapi import HTTPException


@app.post("/upload/", summary='获取文件路径')
async def upload_file(file: UploadFile = File(...)):
    # 只允许上传Excel文件
    allowed_types = [
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",  # .xlsx
        "application/vnd.ms-excel"  # .xls
    ]
    allowed_exts = [".xlsx", ".xls"]

    # 打印上传文件的基本信息
    print(f"收到上传文件: {file.filename}, Content-Type: {file.content_type}")

    # 检查文件类型
    file_ext = os.path.splitext(file.filename)[1].lower()
    if file.content_type not in allowed_types or file_ext not in allowed_exts:
        print(f"文件类型不被允许: {file.filename}, Content-Type: {file.content_type}, 扩展名: {file_ext}")
        raise HTTPException(status_code=500, detail="只允许上传Excel文件（.xlsx或.xls）")

    file_path = os.path.join(FILE_DIR, file.filename)
    with open(file_path, "wb") as buffer:
        file_content = await file.read()
        buffer.write(file_content)
        print(f"文件已保存到: {file_path}，大小: {len(file_content)} 字节")
    return {"file_path": file_path}


@app.get("/getTitle", summary="获取标题")
def getTitle(sheet_name: str, index: int, file_path: str):
    from excel_service import get_title_info
    return get_title_info(sheet_name, index, file_path)
