from typing import Optional
from app.services.graph_client import GraphClient


class OneDriveService:
    def __init__(self, graph_client: GraphClient):
        self.client = graph_client

    async def list_drives(self) -> dict:
        return await self.client.get("/me/drives")

    async def get_drive_root(self, drive_id: Optional[str] = None) -> dict:
        if drive_id:
            return await self.client.get(f"/drives/{drive_id}/root")
        return await self.client.get("/me/drive/root")

    async def get_item(self, item_id: str, drive_id: Optional[str] = None) -> dict:
        if drive_id:
            return await self.client.get(f"/drives/{drive_id}/items/{item_id}")
        return await self.client.get(f"/me/drive/items/{item_id}")

    async def list_children(
        self,
        item_id: str,
        drive_id: Optional[str] = None,
        top: int = 100,
        skip: int = 0,
        order_by: str = "name",
    ) -> dict:
        params = {
            "$top": top,
            "$skip": skip,
            "$orderby": order_by,
        }

        if drive_id:
            endpoint = f"/drives/{drive_id}/items/{item_id}/children"
        else:
            endpoint = f"/me/drive/items/{item_id}/children"

        return await self.client.get(endpoint, params=params)

    async def download_content(self, item_id: str, drive_id: Optional[str] = None) -> bytes:
        if drive_id:
            endpoint = f"/drives/{drive_id}/items/{item_id}/content"
        else:
            endpoint = f"/me/drive/items/{item_id}/content"

        return await self.client.get_raw(endpoint)

    async def upload_content(
        self,
        parent_id: str,
        filename: str,
        content: bytes,
        content_type: str = "application/octet-stream",
        drive_id: Optional[str] = None,
    ) -> dict:
        if drive_id:
            endpoint = f"/drives/{drive_id}/items/{parent_id}:/{filename}:/content"
        else:
            endpoint = f"/me/drive/items/{parent_id}:/{filename}:/content"

        return await self.client.put(endpoint, content, content_type)

    async def delete_item(self, item_id: str, drive_id: Optional[str] = None) -> None:
        if drive_id:
            await self.client.delete(f"/drives/{drive_id}/items/{item_id}")
        else:
            await self.client.delete(f"/me/drive/items/{item_id}")

    async def create_folder(
        self,
        parent_id: str,
        folder_name: str,
        drive_id: Optional[str] = None,
    ) -> dict:
        data = {
            "name": folder_name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "rename",
        }

        if drive_id:
            endpoint = f"/drives/{drive_id}/items/{parent_id}/children"
        else:
            endpoint = f"/me/drive/items/{parent_id}/children"

        return await self.client.post(endpoint, data)

    async def rename_item(
        self,
        item_id: str,
        new_name: str,
        drive_id: Optional[str] = None,
    ) -> dict:
        if drive_id:
            endpoint = f"/drives/{drive_id}/items/{item_id}"
        else:
            endpoint = f"/me/drive/items/{item_id}"

        return await self.client.patch(endpoint, {"name": new_name})

    async def move_item(
        self,
        item_id: str,
        new_parent_id: str,
        new_name: Optional[str] = None,
        drive_id: Optional[str] = None,
    ) -> dict:
        data = {
            "parentReference": {"id": new_parent_id},
        }
        if new_name:
            data["name"] = new_name

        if drive_id:
            endpoint = f"/drives/{drive_id}/items/{item_id}"
        else:
            endpoint = f"/me/drive/items/{item_id}"

        return await self.client.patch(endpoint, data)

    async def search(
        self,
        query: str,
        drive_id: Optional[str] = None,
        top: int = 25,
    ) -> dict:
        params = {"$top": top}

        if drive_id:
            endpoint = f"/drives/{drive_id}/root/search(q='{query}')"
        else:
            endpoint = f"/me/drive/root/search(q='{query}')"

        return await self.client.get(endpoint, params=params)
