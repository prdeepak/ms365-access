"""Contacts router — CRUD for MS365 personal contacts."""

from fastapi import APIRouter, Depends, Query, HTTPException
from typing import Optional

from app.dependencies import get_graph_client, get_current_auth, require_permission
from app.services.graph_client import GraphClient
from app.services.contacts_service import ContactsService
from app.models import Auth
from app.schemas import CreateContactRequest, UpdateContactRequest
from app import audit

router = APIRouter(prefix="/contacts", tags=["contacts"])


def get_contacts_service(
    graph_client: GraphClient = Depends(get_graph_client),
) -> ContactsService:
    return ContactsService(graph_client)


@router.get("/", dependencies=[Depends(require_permission("read:contacts"))])
async def list_contacts(
    top: int = Query(100, ge=1, le=500),
    skip: int = Query(0, ge=0),
    search: Optional[str] = None,
    contacts_service: ContactsService = Depends(get_contacts_service),
):
    result = await contacts_service.list_contacts(top=top, skip=skip, search=search)
    return result.get("value", [])


@router.get("/by-email/{email}", dependencies=[Depends(require_permission("read:contacts"))])
async def search_by_email(
    email: str,
    contacts_service: ContactsService = Depends(get_contacts_service),
):
    result = await contacts_service.search_by_email(email)
    contacts = result.get("value", [])
    if not contacts:
        return {"found": False, "contact": None}
    return {"found": True, "contact": contacts[0]}


@router.get("/{contact_id}", dependencies=[Depends(require_permission("read:contacts"))])
async def get_contact(
    contact_id: str,
    contacts_service: ContactsService = Depends(get_contacts_service),
):
    return await contacts_service.get_contact(contact_id)


@router.post("/", dependencies=[Depends(require_permission("write:contacts"))])
async def create_contact(
    request: CreateContactRequest,
    contacts_service: ContactsService = Depends(get_contacts_service),
    auth: Auth = Depends(get_current_auth),
):
    result = await contacts_service.create_contact(
        name=request.name,
        email=request.email,
        phone=request.phone,
        organization=request.organization,
        title=request.title,
        notes=request.notes,
    )
    audit.log_event(
        "contacts", "create",
        email=auth.email,
        details={"name": request.name, "contact_email": request.email},
    )
    return result


@router.patch("/{contact_id}", dependencies=[Depends(require_permission("write:contacts"))])
async def update_contact(
    contact_id: str,
    request: UpdateContactRequest,
    contacts_service: ContactsService = Depends(get_contacts_service),
    auth: Auth = Depends(get_current_auth),
):
    result = await contacts_service.update_contact(
        contact_id=contact_id,
        name=request.name,
        email=request.email,
        phone=request.phone,
        organization=request.organization,
        title=request.title,
        notes=request.notes,
    )
    changed = [k for k, v in request.model_dump(exclude_none=True).items()]
    audit.log_event(
        "contacts", "update",
        email=auth.email,
        details={"contact_id": contact_id, "changed_fields": changed},
    )
    return result


@router.delete("/{contact_id}", dependencies=[Depends(require_permission("write:contacts"))])
async def delete_contact(
    contact_id: str,
    contacts_service: ContactsService = Depends(get_contacts_service),
    auth: Auth = Depends(get_current_auth),
):
    await contacts_service.delete_contact(contact_id)
    audit.log_event(
        "contacts", "delete",
        email=auth.email,
        details={"contact_id": contact_id},
    )
    return {"success": True, "deleted": contact_id}
