from typing import Optional
from urllib.parse import urlparse, unquote

from app.services.graph_client import GraphClient


class SharePointService:
    def __init__(self, graph_client: GraphClient):
        self.client = graph_client

    async def resolve_site(self, host_and_path: str) -> dict:
        """Resolve a SharePoint site by hostname and path.

        Args:
            host_and_path: e.g. "contoso.sharepoint.com:/sites/Finance"
                           or "contoso.sharepoint.com/sites/Finance"
        """
        # Normalize: ensure colon separator between host and site path
        if ":/" not in host_and_path:
            # Split on first /sites/ or /teams/ occurrence
            for prefix in ("/sites/", "/teams/"):
                idx = host_and_path.find(prefix.lstrip("/"))
                if idx > 0:
                    hostname = host_and_path[:idx].rstrip("/")
                    site_path = host_and_path[idx:]
                    if not site_path.startswith("/"):
                        site_path = "/" + site_path
                    host_and_path = f"{hostname}:{site_path}"
                    break

        return await self.client.get(f"/sites/{host_and_path}")

    async def list_drives(self, site_id: str) -> dict:
        """List document libraries (drives) for a site."""
        return await self.client.get(f"/sites/{site_id}/drives")

    async def get_drive(self, site_id: str, drive_id: str) -> dict:
        """Get a specific drive by ID."""
        return await self.client.get(f"/sites/{site_id}/drives/{drive_id}")

    async def list_children(
        self,
        drive_id: str,
        item_id: str = "root",
        top: int = 100,
        order_by: str = "name",
    ) -> dict:
        """List children of a folder in a SharePoint drive."""
        params = {
            "$top": top,
            "$orderby": order_by,
        }
        return await self.client.get(
            f"/drives/{drive_id}/items/{item_id}/children", params=params
        )

    async def search(
        self,
        drive_id: str,
        query: str,
        top: int = 25,
    ) -> dict:
        """Search within a SharePoint drive."""
        params = {"$top": top}
        return await self.client.get(
            f"/drives/{drive_id}/root/search(q='{query}')", params=params
        )

    async def get_item(self, drive_id: str, item_id: str) -> dict:
        """Get item metadata."""
        return await self.client.get(f"/drives/{drive_id}/items/{item_id}")

    async def download_content(
        self,
        drive_id: str,
        item_id: str,
        format: Optional[str] = None,
    ) -> bytes:
        """Download file content, optionally converting format.

        Args:
            format: Optional conversion format (e.g. "pdf").
        """
        endpoint = f"/drives/{drive_id}/items/{item_id}/content"
        if format:
            endpoint += f"?format={format}"
        return await self.client.get_raw(endpoint)

    async def resolve_sharepoint_url(self, url: str) -> dict:
        """Parse a SharePoint sharing URL and resolve to item metadata.

        Handles URLs like:
          https://contoso.sharepoint.com/:w:/s/SiteName/EaBC123...
          https://contoso.sharepoint.com/sites/SiteName/Shared Documents/file.docx

        Returns dict with site, drive_id, item_id, and item metadata.
        """
        parsed = urlparse(url)
        hostname = parsed.hostname

        # Try the shares API first (works for sharing links)
        # Encode the URL as a sharing token
        import base64
        encoded = base64.urlsafe_b64encode(url.encode()).decode().rstrip("=")
        sharing_token = f"u!{encoded}"

        try:
            result = await self.client.get(
                f"/shares/{sharing_token}/driveItem",
                params={"$expand": "parentReference"},
            )
            return {
                "item": result,
                "item_id": result.get("id"),
                "drive_id": result.get("parentReference", {}).get("driveId"),
                "site_id": result.get("parentReference", {}).get("siteId"),
            }
        except Exception:
            pass

        # Fallback: try to parse the URL structure directly
        path = unquote(parsed.path)

        # Extract site path (e.g., /sites/Finance)
        site_path = None
        for prefix in ("/sites/", "/teams/"):
            if prefix in path:
                idx = path.index(prefix)
                # Get site name (next path segment)
                rest = path[idx + len(prefix):]
                site_name = rest.split("/")[0]
                site_path = f"{prefix}{site_name}"
                doc_path = rest[len(site_name):]
                break

        if not site_path:
            raise ValueError(f"Cannot parse SharePoint site from URL: {url}")

        # Resolve the site
        site = await self.resolve_site(f"{hostname}{site_path}")
        site_id = site["id"]

        # Get drives and try to find the item by path
        drives_result = await self.list_drives(site_id)
        drives = drives_result.get("value", [])

        # Try each drive to find the item
        if doc_path:
            # Strip common prefixes like /Shared Documents
            for drive in drives:
                drive_id = drive["id"]
                try:
                    item = await self.client.get(
                        f"/drives/{drive_id}/root:{doc_path}"
                    )
                    return {
                        "item": item,
                        "item_id": item.get("id"),
                        "drive_id": drive_id,
                        "site_id": site_id,
                    }
                except Exception:
                    continue

        return {
            "site": site,
            "site_id": site_id,
            "drives": drives,
            "error": "Could not resolve to a specific item",
        }
