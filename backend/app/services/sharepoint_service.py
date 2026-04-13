import logging
from typing import Optional
from urllib.parse import urlparse, unquote, quote

from app.services.graph_client import GraphClient

logger = logging.getLogger(__name__)


class SharePointService:
    def __init__(self, graph_client: GraphClient):
        self.client = graph_client

    def _drive_path(self, site_id: str) -> str:
        """Build the Graph API drive prefix using site context.

        Uses /sites/{site_id}/drive which avoids b!-prefixed drive IDs
        that cause 400 errors with httpx URL encoding.
        """
        return f"/sites/{site_id}/drive"

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
        site_id: str,
        item_id: str = "root",
        top: int = 100,
        order_by: str = "name",
    ) -> dict:
        """List children of a folder in a SharePoint drive."""
        params = {
            "$top": top,
            "$orderby": order_by,
        }
        drive = self._drive_path(site_id)
        return await self.client.get(
            f"{drive}/items/{item_id}/children", params=params
        )

    async def search(
        self,
        site_id: str,
        query: str,
        top: int = 25,
    ) -> dict:
        """Search within a SharePoint site's default drive."""
        safe_query = quote(query, safe='')
        params = {"$top": top}
        drive = self._drive_path(site_id)
        return await self.client.get(
            f"{drive}/root/search(q='{safe_query}')", params=params
        )

    async def get_item(self, site_id: str, item_id: str) -> dict:
        """Get item metadata."""
        drive = self._drive_path(site_id)
        return await self.client.get(f"{drive}/items/{item_id}")

    async def rename_item(self, site_id: str, item_id: str, new_name: str) -> dict:
        """Rename a file or folder in a SharePoint drive."""
        drive = self._drive_path(site_id)
        return await self.client.patch(f"{drive}/items/{item_id}", {"name": new_name})

    async def move_item(
        self, site_id: str, item_id: str, destination_folder_id: str
    ) -> dict:
        """Move a file or folder to a different folder in a SharePoint drive."""
        drive = self._drive_path(site_id)
        return await self.client.patch(
            f"{drive}/items/{item_id}",
            {"parentReference": {"id": destination_folder_id}},
        )

    async def download_content(
        self,
        site_id: str,
        item_id: str,
        format: Optional[str] = None,
    ) -> bytes:
        """Download file content, optionally converting format.

        Args:
            format: Optional conversion format (e.g. "pdf").
        """
        drive = self._drive_path(site_id)
        endpoint = f"{drive}/items/{item_id}/content"
        if format:
            endpoint += f"?format={format}"
        return await self.client.get_raw(endpoint)

    async def upload_content(
        self,
        site_id: str,
        parent_id: str,
        filename: str,
        content: bytes,
        content_type: str = "application/octet-stream",
    ) -> dict:
        """Upload a file to a SharePoint site's default drive.

        Uses /sites/{site_id}/drive path to avoid b!-prefixed drive IDs.
        """
        drive = self._drive_path(site_id)
        endpoint = f"{drive}/items/{parent_id}:/{filename}:/content"
        return await self.client.put(endpoint, content, content_type)

    async def replace_content(
        self,
        site_id: str,
        item_id: str,
        content: bytes,
        content_type: str = "application/octet-stream",
    ) -> dict:
        """Replace the content of an existing file in a SharePoint drive (keeps the same item ID)."""
        drive = self._drive_path(site_id)
        endpoint = f"{drive}/items/{item_id}/content"
        return await self.client.put(endpoint, content, content_type)

    async def resolve_sharepoint_url(self, url: str) -> dict:
        """Parse a SharePoint sharing URL and resolve to item metadata.

        Handles URLs like:
          https://contoso.sharepoint.com/:w:/s/SiteName/EaBC123...
          https://contoso.sharepoint.com/sites/SiteName/Shared Documents/file.docx

        Returns dict with site_id, item_id, and item metadata.
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
                extra_headers={"Prefer": "redeemSharingLinkIfNecessary"},
            )
            return {
                "item": result,
                "item_id": result.get("id"),
                "site_id": result.get("parentReference", {}).get("siteId"),
            }
        except Exception as e:
            detail = ""
            if hasattr(e, "response"):
                try:
                    detail = e.response.json()
                except Exception:
                    detail = getattr(e.response, "text", "")
            logger.warning("Shares API failed for %s: %s (detail: %s)", url, e, detail)
            # Fall through to URL-parsing fallback

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

        # Try to find the item by path using the site's default drive
        if doc_path:
            # Strip the library name prefix if present (e.g. "/Shared Documents/...")
            drives_result = await self.list_drives(site_id)
            drives = drives_result.get("value", [])
            for drive in drives:
                drive_web_url = drive.get("webUrl", "")
                if drive_web_url:
                    lib_suffix = urlparse(drive_web_url).path.split("/")[-1]
                    lib_prefix = f"/{unquote(lib_suffix)}"
                    if doc_path.startswith(lib_prefix):
                        doc_path = doc_path[len(lib_prefix):] or "/"
                        break

            if doc_path and doc_path != "/":
                try:
                    drive = self._drive_path(site_id)
                    item = await self.client.get(f"{drive}/root:{doc_path}")
                    return {
                        "item": item,
                        "item_id": item.get("id"),
                        "site_id": site_id,
                    }
                except Exception:
                    pass

        raise ValueError(f"Could not resolve URL to a specific item: {url}")
