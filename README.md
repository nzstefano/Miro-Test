# Project Setup & Usage

This project is a backend service built with **Node.js** and integrates with a **Python script** to generate PPTX files.  
The Node.js server handles API requests, passes data to Python, and returns the generated file path.

---

## 📦 Dependencies

Before running the project, install the following:

- [Node.js](https://nodejs.org/en/download/) v16 or higher
- [Yarn](https://classic.yarnpkg.com/en/docs/install)
- [Python 3](https://www.python.org/downloads/) (v3.8 or higher)
- [Git](https://git-scm.com/downloads)
- [Postman](https://www.postman.com/downloads/) or `curl` (for sending test requests)

Node.js main packages (already listed in `package.json`):

- `express` → Web server
- `axios` → Handle API calls / communication
- `multer` or `fs` → File handling
- `child_process` → Run Python script from Node.js

Python main packages (already listed in `requirements.txt`):

- `python-pptx` → Generate PPTX files
- `json` (builtin) → Parse input JSON

---

## 🚀 Installation

1. **Clone the repository**
   ```bash
   git clone <repo-url>
   cd <project-folder>
   ```
   2
   . **Install Node.js dependencies**
   ```bash
   yarn install
   ```
2. **Install Python dependencies**
   ```bash
   pip install -r converter/requirements.txt
   ```

## ▶️ Running the Server

**Start the Node.js server:**

```bash
yarn start
```

By default, the server will run at:

http://localhost:3000

## 📮 Sending a Request

You can send requests either using Postman or curl.

Example JSON Input

```json
{
  "content": {
    "widgets": [
      {
        "id": "3458764639651442143",
        "canvasedObjectData": {
          "widgetId": null,
          "type": "slidecontainer",
          "json": "{\"rotation\":{\"rotation\":0.0},\"scale\":{\"scale\":1.0},\"relativeRotation\":0,\"relativeScale\":1,\"direction\":2,\"padding\":57.02}"
        }
      }
    ]
  }
}
```

### Example with Postman

1. Open Postman

2. Create a POST request to:

http://localhost:3000/convert

3. Go to Body → raw → JSON

4. Paste the example JSON input

5. Send the request

If successful, you will receive a JSON response with the file path.

### Example with curl

```bash
curl -X POST http://localhost:3000/convert \
 -H "Content-Type: application/json" \
 -d '{
"content": {
"widgets": [
{
"id": "3458764639651442143",
"canvasedObjectData": {
"widgetId": null,
"type": "slidecontainer",
"json": "{\"rotation\":{\"rotation\":0.0},\"scale\":{\"scale\":1.0}}"
}
}
]
}
}'
```

📂 Generated Files

All generated PPTX files are saved inside the folder:

```bash
/output
```

Each file is named with a unique timestamp, for example:

```bash
/output/generated_20250915_123456.pptx
```

You can open this PPTX file directly with Microsoft PowerPoint or Google Slides.

✅ Quick Test

Run the server:

```bash
yarn start
```

output/generated_20250915_123456.pptx

You can open this PPTX file directly with Microsoft PowerPoint or Google Slides.

### ✅ Quick Test

1. Run the server:

```bash
yarn start
```

2. Send a request (Postman or curl).

```bash
curl -X POST http://localhost:3000/convert \
 -H "Content-Type: application/json" \
 -d '{
  "content": {
    "widgets": [
      {
        "id": "3458764639651442143",
        "canvasedObjectData": {
          "widgetId": null,
          "type": "slidecontainer",
          "json": "{\"rotation\":{\"rotation\":0.0},\"scale\":{\"scale\":1.0},\"relativeRotation\":0,\"relativeScale\":1,\"direction\":2,\"padding\":57.02}"
        }
      }
    ]
  }
}'
```

3. Check the /output folder for the generated .pptx file.
