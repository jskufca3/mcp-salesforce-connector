# /// script
# dependencies = [
#   "mcp",
#   "simple-salesforce",
#   "python-dotenv",
#   "python-docx",
#   "requests",
#   "lxml",
#   "pypdf"
# ]
# ///
import asyncio
import json
from typing import Any, Optional
import os
import io
import zipfile
from dotenv import load_dotenv

import requests
from lxml import etree
from docx import Document
from pypdf import PdfReader

from simple_salesforce import Salesforce
from simple_salesforce.exceptions import SalesforceError

import mcp.types as types
from mcp.server import Server, NotificationOptions
from mcp.server.models import InitializationOptions
import mcp.server.stdio

class SalesforceClient:
    """Handles Salesforce operations and caching."""
    
    def __init__(self):
        self.sf: Optional[Salesforce] = None
        self.sobjects_cache: dict[str, Any] = {}

    def connect(self) -> bool:
        """Establishes connection to Salesforce using environment variables.
        
        Returns:
            bool: True if connection successful, False otherwise
        """
        try:
            access_token = os.getenv('SALESFORCE_ACCESS_TOKEN')
            instance_url = os.getenv('SALESFORCE_INSTANCE_URL')
            domain = os.getenv('SALESFORCE_DOMAIN')
            if access_token and instance_url:
                self.sf = Salesforce(
                    instance_url=instance_url,
                    session_id=access_token,
                    domain=domain
                )
                return True
            
            self.sf = Salesforce(
                username=os.getenv('SALESFORCE_USERNAME'),
                password=os.getenv('SALESFORCE_PASSWORD'),
                security_token=os.getenv('SALESFORCE_SECURITY_TOKEN'),
                domain=domain
            )
            return True
        except Exception as e:
            print(f"Salesforce connection failed: {str(e)}")
            return False
    
    def get_object_fields(self, object_name: str) -> str:
        """Retrieves field Names, labels and types for a specific Salesforce object.

        Args:
            object_name (str): The name of the Salesforce object.

        Returns:
            str: JSON representation of the object fields.
        """
        if not self.sf:
            raise ValueError("Salesforce connection not established.")
        if object_name not in self.sobjects_cache:
            sf_object = getattr(self.sf, object_name)
            fields = sf_object.describe()['fields']
            filtered_fields = []
            for field in fields:
                filtered_fields.append({
                    'label': field['label'],
                    'name': field['name'],
                    'updateable': field['updateable'],
                    'type': field['type'],
                    'length': field['length'],
                    'picklistValues': field['picklistValues']
                })
            self.sobjects_cache[object_name] = filtered_fields
            
        return json.dumps(self.sobjects_cache[object_name], indent=2)

    def download_content_version_text(self, content_version_id: str) -> str:
        """Downloads a ContentVersion file and extracts text content.

        Supports .docx files. Returns raw bytes as base64 for other types.

        Args:
            content_version_id (str): The Salesforce ContentVersion ID.

        Returns:
            str: Extracted text content.
        """
        if not self.sf:
            raise ValueError("Salesforce connection not established.")

        # Build the VersionData URL and fetch with session auth
        instance_url = self.sf.base_url.split('/services')[0]
        url = f"{instance_url}/services/data/v59.0/sobjects/ContentVersion/{content_version_id}/VersionData"
        headers = {
            "Authorization": f"Bearer {self.sf.session_id}",
            "Content-Type": "application/octet-stream",
        }
        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()

        # Determine file type from ContentVersion metadata
        cv_record = self.sf.query(
            f"SELECT FileExtension FROM ContentVersion WHERE Id = '{content_version_id}'"
        )
        file_ext = ""
        if cv_record['records']:
            file_ext = (cv_record['records'][0].get('FileExtension') or "").lower()

        if file_ext == "pdf":
            reader = PdfReader(io.BytesIO(response.content))
            pages = []
            for page in reader.pages:
                text = page.extract_text()
                if text and text.strip():
                    pages.append(text.strip())
            return "\n\n".join(pages)

        elif file_ext in ("docx",):
            # Use lxml to do a full XML traversal — captures headers, footers,
            # text boxes, table cells, and nested content that python-docx misses.
            WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            W_P = f"{{{WORD_NS}}}p"
            W_T = f"{{{WORD_NS}}}t"

            all_parts = []
            with zipfile.ZipFile(io.BytesIO(response.content)) as z:
                # Collect document body first, then headers/footers
                targets = (
                    ["word/document.xml"]
                    + sorted(n for n in z.namelist()
                             if n.startswith("word/header") and n.endswith(".xml"))
                    + sorted(n for n in z.namelist()
                             if n.startswith("word/footer") and n.endswith(".xml"))
                )
                for xml_name in targets:
                    if xml_name not in z.namelist():
                        continue
                    tree = etree.fromstring(z.read(xml_name))
                    for para in tree.iter(W_P):
                        text = "".join(t.text or "" for t in para.iter(W_T))
                        if text.strip():
                            all_parts.append(text)
            return "\n".join(all_parts)
        else:
            # Unsupported binary — return a message with size info
            return f"[Binary file ({file_ext}, {len(response.content)} bytes) — text extraction not supported for this format]"


# Create a server instance
server = Server("salesforce-mcp")

# Load environment variables
load_dotenv()

# Configure with Salesforce credentials from environment variables
sf_client = SalesforceClient()
if not sf_client.connect():
    print("Failed to initialize Salesforce connection")

# Add tool capabilities to run SOQL queries
@server.list_tools()
async def handle_list_tools() -> list[types.Tool]:
    """
    List available tools.
    Each tool specifies its arguments using JSON Schema validation.
    """
    return [
        types.Tool(
            name="run_soql_query",
            description="Executes a SOQL query against Salesforce",
            inputSchema={
                "type": "object",
                "properties": {
                    "query": {
                        "type": "string",
                        "description": "The SOQL query to execute",
                    },
                },
                "required": ["query"],
            },
        ),
        types.Tool(
            name="run_sosl_search",
            description="Executes a SOSL search against Salesforce",
            inputSchema={
                "type": "object",
                "properties": {
                    "search": {
                        "type": "string",
                        "description": "The SOSL search to execute (e.g., 'FIND {John Smith} IN ALL FIELDS')",
                    },
                },
                "required": ["search"],
            },
        ),
        types.Tool(
            name="get_object_fields",
            description="Retrieves field Names, labels and types for a specific Salesforce object",
            inputSchema={
                "type": "object",
                "properties": {
                    "object_name": {
                        "type": "string",
                        "description": "The name of the Salesforce object (e.g., 'Account', 'Contact')",
                    },
                },
                "required": ["object_name"],
            },
        ),
        types.Tool(
            name="get_record",
            description="Retrieves a specific record by ID",
            inputSchema={
                "type": "object",
                "properties": {
                    "object_name": {
                        "type": "string",
                        "description": "The name of the Salesforce object (e.g., 'Account', 'Contact')",
                    },
                    "record_id": {
                        "type": "string",
                        "description": "The ID of the record to retrieve",
                    },
                },
                "required": ["object_name", "record_id"],
            },
        ),
        types.Tool(
            name="create_record",
            description="Creates a new record",
            inputSchema={
                "type": "object",
                "properties": {
                    "object_name": {
                        "type": "string",
                        "description": "The name of the Salesforce object (e.g., 'Account', 'Contact')",
                    },
                    "data": {
                        "type": "object",
                        "description": "The data for the new record",
                        "properties": {},
                        "additionalProperties": True,
                    },
                },
                "required": ["object_name", "data"],
            },
        ),
        types.Tool(
            name="update_record",
            description="Updates an existing record",
            inputSchema={
                "type": "object",
                "properties": {
                    "object_name": {
                        "type": "string",
                        "description": "The name of the Salesforce object (e.g., 'Account', 'Contact')",
                    },
                    "record_id": {
                        "type": "string",
                        "description": "The ID of the record to update",
                    },
                    "data": {
                        "type": "object",
                        "description": "The updated data for the record",
                        "properties": {},
                        "additionalProperties": True,
                    },
                },
                "required": ["object_name", "record_id", "data"],
            },
        ),
        types.Tool(
            name="delete_record",
            description="Deletes a record",
            inputSchema={
                "type": "object",
                "properties": {
                    "object_name": {
                        "type": "string",
                        "description": "The name of the Salesforce object (e.g., 'Account', 'Contact')",
                    },
                    "record_id": {
                        "type": "string",
                        "description": "The ID of the record to delete",
                    },
                },
                "required": ["object_name", "record_id"],
            },
        ),
        types.Tool(
            name="tooling_execute",
            description="Executes a Tooling API request",
            inputSchema={
                "type": "object",
                "properties": {
                    "action": {
                        "type": "string",
                        "description": "The Tooling API endpoint to call (e.g., 'sobjects/ApexClass')",
                    },
                    "method": {
                        "type": "string",
                        "description": "The HTTP method (default: 'GET')",
                        "enum": ["GET", "POST", "PATCH", "DELETE"],
                        "default": "GET",
                    },
                    "data": {
                        "type": "object",
                        "description": "Data for POST/PATCH requests",
                        "properties": {},
                        "additionalProperties": True,
                    },
                },
                "required": ["action"],
            },
        ),
        types.Tool(
            name="apex_execute",
            description="Executes an Apex REST request",
            inputSchema={
                "type": "object",
                "properties": {
                    "action": {
                        "type": "string",
                        "description": "The Apex REST endpoint to call (e.g., '/MyApexClass')",
                    },
                    "method": {
                        "type": "string",
                        "description": "The HTTP method (default: 'GET')",
                        "enum": ["GET", "POST", "PATCH", "DELETE"],
                        "default": "GET",
                    },
                    "data": {
                        "type": "object",
                        "description": "Data for POST/PATCH requests",
                        "properties": {},
                        "additionalProperties": True,
                    },
                },
                "required": ["action"],
            },
        ),
        types.Tool(
            name="restful",
            description="Makes a direct REST API call to Salesforce",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "The path of the REST API endpoint (e.g., 'sobjects/Account/describe')",
                    },
                    "method": {
                        "type": "string",
                        "description": "The HTTP method (default: 'GET')",
                        "enum": ["GET", "POST", "PATCH", "DELETE"],
                        "default": "GET",
                    },
                    "params": {
                        "type": "object",
                        "description": "Query parameters for the request",
                        "properties": {},
                        "additionalProperties": True,
                    },
                    "data": {
                        "type": "object",
                        "description": "Data for POST/PATCH requests",
                        "properties": {},
                        "additionalProperties": True,
                    },
                },
                "required": ["path"],
            },
        ),
        types.Tool(
            name="download_content_version",
            description="Downloads a Salesforce ContentVersion file and extracts its text content. Supports .docx files. Use this to read Word documents attached to records.",
            inputSchema={
                "type": "object",
                "properties": {
                    "content_version_id": {
                        "type": "string",
                        "description": "The Salesforce ContentVersion ID (starts with '068')",
                    },
                },
                "required": ["content_version_id"],
            },
        ),
    ]

@server.call_tool()
async def handle_call_tool(name: str, arguments: dict[str, str]) -> list[types.TextContent]:
    if name == "run_soql_query":
        query = arguments.get("query")
        if not query:
            raise ValueError("Missing 'query' argument")

        results = sf_client.sf.query_all(query)
        return [
            types.TextContent(
                type="text",
                text=f"SOQL Query Results (JSON):\n{json.dumps(results, indent=2)}",
            )
        ]
    elif name == "run_sosl_search":
        search = arguments.get("search")
        if not search:
            raise ValueError("Missing 'search' argument")

        results = sf_client.sf.search(search)
        return [
            types.TextContent(
                type="text",
                text=f"SOSL Search Results (JSON):\n{json.dumps(results, indent=2)}",
            )
        ]
    elif name == "get_object_fields":
        object_name = arguments.get("object_name")
        if not object_name:
            raise ValueError("Missing 'object_name' argument")
        if not sf_client.sf:
            raise ValueError("Salesforce connection not established.")
        results = sf_client.get_object_fields(object_name)
        return [
            types.TextContent(
                type="text",
                text=f"{object_name} Metadata (JSON):\n{results}",
            )
        ]
    elif name == "get_record":
        object_name = arguments.get("object_name")
        record_id = arguments.get("record_id")
        if not object_name or not record_id:
            raise ValueError("Missing 'object_name' or 'record_id' argument")
        if not sf_client.sf:
            raise ValueError("Salesforce connection not established.")
        sf_object = getattr(sf_client.sf, object_name)
        results = sf_object.get(record_id)
        return [
            types.TextContent(
                type="text",
                text=f"{object_name} Record (JSON):\n{json.dumps(results, indent=2)}",
            )
        ]
    elif name == "create_record":
        object_name = arguments.get("object_name")
        data = arguments.get("data")
        if not object_name or not data:
            raise ValueError("Missing 'object_name' or 'data' argument")
        if not sf_client.sf:
            raise ValueError("Salesforce connection not established.")
        sf_object = getattr(sf_client.sf, object_name)
        results = sf_object.create(data)
        return [
            types.TextContent(
                type="text",
                text=f"Create {object_name} Record Result (JSON):\n{json.dumps(results, indent=2)}",
            )
        ]
    elif name == "update_record":
        object_name = arguments.get("object_name")
        record_id = arguments.get("record_id")
        data = arguments.get("data")
        if not object_name or not record_id or not data:
            raise ValueError("Missing 'object_name', 'record_id', or 'data' argument")
        if not sf_client.sf:
            raise ValueError("Salesforce connection not established.")
        sf_object = getattr(sf_client.sf, object_name)
        results = sf_object.update(record_id, data)
        return [
            types.TextContent(
                type="text",
                text=f"Update {object_name} Record Result: {results}",
            )
        ]
    elif name == "delete_record":
        object_name = arguments.get("object_name")
        record_id = arguments.get("record_id")
        if not object_name or not record_id:
            raise ValueError("Missing 'object_name' or 'record_id' argument")
        if not sf_client.sf:
            raise ValueError("Salesforce connection not established.")
        sf_object = getattr(sf_client.sf, object_name)
        results = sf_object.delete(record_id)
        return [
            types.TextContent(
                type="text",
                text=f"Delete {object_name} Record Result: {results}",
            )
        ]
    elif name == "tooling_execute":
        action = arguments.get("action")
        method = arguments.get("method", "GET")
        data = arguments.get("data")

        if not action:
            raise ValueError("Missing 'action' argument")
        if not sf_client.sf:
            raise ValueError("Salesforce connection not established.")

        results = sf_client.sf.toolingexecute(action, method=method, data=data)
        return [
            types.TextContent(
                type="text",
                text=f"Tooling Execute Result (JSON):\n{json.dumps(results, indent=2)}",
            )
        ]
    elif name == "apex_execute":
        action = arguments.get("action")
        method = arguments.get("method", "GET")
        data = arguments.get("data")

        if not action:
            raise ValueError("Missing 'action' argument")
        if not sf_client.sf:
            raise ValueError("Salesforce connection not established.")

        results = sf_client.sf.apexecute(action, method=method, data=data)
        return [
            types.TextContent(
                type="text",
                text=f"Apex Execute Result (JSON):\n{json.dumps(results, indent=2)}",
            )
        ]
    elif name == "restful":
        path = arguments.get("path")
        method = arguments.get("method", "GET")
        params = arguments.get("params")
        data = arguments.get("data")

        if not path:
            raise ValueError("Missing 'path' argument")
        if not sf_client.sf:
            raise ValueError("Salesforce connection not established.")

        results = sf_client.sf.restful(path, method=method, params=params, json=data)
        return [
            types.TextContent(
                type="text",
                text=f"RESTful API Call Result (JSON):\n{json.dumps(results, indent=2)}",
            )
        ]
    elif name == "download_content_version":
        content_version_id = arguments.get("content_version_id")
        if not content_version_id:
            raise ValueError("Missing 'content_version_id' argument")
        if not sf_client.sf:
            raise ValueError("Salesforce connection not established.")

        text = sf_client.download_content_version_text(content_version_id)
        return [
            types.TextContent(
                type="text",
                text=f"ContentVersion {content_version_id} Text Content:\n\n{text}",
            )
        ]
    raise ValueError(f"Unknown tool: {name}")


async def run():
    async with mcp.server.stdio.stdio_server() as (read, write):
        await server.run(
            read,
            write,
            InitializationOptions(
                server_name="salesforce-mcp",
                server_version="0.1.6",
                capabilities=server.get_capabilities(
                    notification_options=NotificationOptions(),
                    experimental_capabilities={},
                ),
            ),
        )

if __name__ == "__main__":
    asyncio.run(run())
