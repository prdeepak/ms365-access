"""MS365 Contacts service — CRUD operations via MS Graph /me/contacts."""

from typing import Optional
from app.services.graph_client import GraphClient


class ContactsService:
    def __init__(self, graph_client: GraphClient):
        self.client = graph_client

    async def list_contacts(
        self,
        top: int = 100,
        skip: int = 0,
        search: Optional[str] = None,
    ) -> dict:
        params = {
            "$top": top,
            "$skip": skip,
            "$orderby": "displayName",
        }
        if search:
            params["$search"] = f'"{search}"'
            # Graph search requires ConsistencyLevel header
            params["$count"] = "true"

        return await self.client.get("/me/contacts", params=params)

    async def get_contact(self, contact_id: str) -> dict:
        return await self.client.get(f"/me/contacts/{contact_id}")

    async def create_contact(
        self,
        name: str,
        email: Optional[str] = None,
        phone: Optional[str] = None,
        organization: Optional[str] = None,
        title: Optional[str] = None,
        notes: Optional[str] = None,
    ) -> dict:
        # Split name into given/surname
        parts = name.strip().split(None, 1)
        given_name = parts[0] if parts else name
        surname = parts[1] if len(parts) > 1 else ""

        contact_data = {
            "givenName": given_name,
            "surname": surname,
            "displayName": name,
        }

        if email:
            contact_data["emailAddresses"] = [{"address": email}]

        if phone:
            contact_data["mobilePhone"] = phone

        if organization:
            contact_data["companyName"] = organization

        if title:
            contact_data["jobTitle"] = title

        if notes:
            contact_data["personalNotes"] = notes

        return await self.client.post("/me/contacts", contact_data)

    async def update_contact(
        self,
        contact_id: str,
        name: Optional[str] = None,
        email: Optional[str] = None,
        phone: Optional[str] = None,
        organization: Optional[str] = None,
        title: Optional[str] = None,
        notes: Optional[str] = None,
    ) -> dict:
        contact_data = {}

        if name is not None:
            parts = name.strip().split(None, 1)
            contact_data["givenName"] = parts[0] if parts else name
            contact_data["surname"] = parts[1] if len(parts) > 1 else ""
            contact_data["displayName"] = name

        if email is not None:
            contact_data["emailAddresses"] = [{"address": email}]

        if phone is not None:
            contact_data["mobilePhone"] = phone

        if organization is not None:
            contact_data["companyName"] = organization

        if title is not None:
            contact_data["jobTitle"] = title

        if notes is not None:
            contact_data["personalNotes"] = notes

        return await self.client.patch(f"/me/contacts/{contact_id}", contact_data)

    async def delete_contact(self, contact_id: str) -> None:
        await self.client.delete(f"/me/contacts/{contact_id}")

    async def search_by_email(self, email: str) -> dict:
        """Search contacts by exact email address."""
        filter_q = f"emailAddresses/any(e:e/address eq '{email}')"
        return await self.client.get("/me/contacts", params={"$filter": filter_q})
