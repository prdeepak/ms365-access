from fastapi import APIRouter, Depends, Query
from typing import Optional
from datetime import datetime
from zoneinfo import ZoneInfo

from app.dependencies import get_graph_client, get_current_auth, require_permission
from app.services.graph_client import GraphClient
from app.services.calendar_service import CalendarService
from app.models import Auth
from app.schemas import (
    CreateEventRequest,
    UpdateEventRequest,
    RespondEventRequest,
)
from app import audit
from app.config import get_settings

router = APIRouter(prefix="/calendar", tags=["calendar"])


def convert_event_to_local_tz(event: dict, local_tz: str) -> dict:
    """Convert event start/end times from UTC to local timezone."""
    tz = ZoneInfo(local_tz)

    for field in ("start", "end"):
        if field in event and "dateTime" in event[field]:
            # Parse the UTC datetime (Graph API returns ISO format)
            dt_str = event[field]["dateTime"]
            # Handle both with and without Z suffix
            if dt_str.endswith("Z"):
                dt_str = dt_str[:-1]
            try:
                utc_dt = datetime.fromisoformat(dt_str).replace(tzinfo=ZoneInfo("UTC"))
                local_dt = utc_dt.astimezone(tz)
                event[field]["dateTime"] = local_dt.strftime("%Y-%m-%dT%H:%M:%S")
                event[field]["timeZone"] = local_tz
            except (ValueError, KeyError):
                pass  # Leave unchanged if parsing fails

    return event


def get_calendar_service(graph_client: GraphClient = Depends(get_graph_client)) -> CalendarService:
    return CalendarService(graph_client)


@router.get("/calendars", dependencies=[Depends(require_permission("read:calendar"))])
async def list_calendars(calendar_service: CalendarService = Depends(get_calendar_service)):
    result = await calendar_service.list_calendars()
    return result.get("value", [])


@router.get("/events", dependencies=[Depends(require_permission("read:calendar"))])
async def list_events(
    calendar_id: Optional[str] = None,
    top: int = Query(25, ge=1, le=100),
    skip: int = Query(0, ge=0),
    order_by: str = "start/dateTime",
    filter: Optional[str] = None,
    calendar_service: CalendarService = Depends(get_calendar_service),
):
    settings = get_settings()
    result = await calendar_service.list_events(
        calendar_id=calendar_id,
        top=top,
        skip=skip,
        order_by=order_by,
        filter_query=filter,
    )
    items = [convert_event_to_local_tz(e, settings.local_timezone) for e in result.get("value", [])]
    return items


@router.get("/view", dependencies=[Depends(require_permission("read:calendar"))])
async def get_calendar_view(
    start_datetime: datetime,
    end_datetime: datetime,
    calendar_id: Optional[str] = None,
    top: int = Query(100, ge=1, le=500),
    calendar_service: CalendarService = Depends(get_calendar_service),
):
    settings = get_settings()
    result = await calendar_service.get_calendar_view(
        start_datetime=start_datetime,
        end_datetime=end_datetime,
        calendar_id=calendar_id,
        top=top,
    )
    items = [convert_event_to_local_tz(e, settings.local_timezone) for e in result.get("value", [])]
    return items


@router.get("/events/{event_id}", dependencies=[Depends(require_permission("read:calendar"))])
async def get_event(
    event_id: str,
    calendar_service: CalendarService = Depends(get_calendar_service),
):
    settings = get_settings()
    event = await calendar_service.get_event(event_id)
    return convert_event_to_local_tz(event, settings.local_timezone)


@router.post("/events", dependencies=[Depends(require_permission("write:calendar"))])
async def create_event(
    request: CreateEventRequest,
    calendar_id: Optional[str] = None,
    calendar_service: CalendarService = Depends(get_calendar_service),
    auth: Auth = Depends(get_current_auth),
):
    result = await calendar_service.create_event(
        subject=request.subject,
        start_datetime=request.start_datetime,
        end_datetime=request.end_datetime,
        time_zone=request.time_zone,
        body=request.body,
        body_type=request.body_type,
        location=request.location,
        attendees=request.attendees,
        is_all_day=request.is_all_day,
        is_online_meeting=request.is_online_meeting,
        recurrence=request.recurrence,
        reminder_minutes=request.reminder_minutes,
        show_as=request.show_as,
        importance=request.importance,
        calendar_id=calendar_id,
    )
    audit.log_calendar_create(
        auth.email,
        request.subject,
        request.start_datetime.isoformat(),
        request.attendees,
    )
    return result


@router.patch("/events/{event_id}", dependencies=[Depends(require_permission("write:calendar"))])
async def update_event(
    event_id: str,
    request: UpdateEventRequest,
    calendar_service: CalendarService = Depends(get_calendar_service),
    auth: Auth = Depends(get_current_auth),
):
    result = await calendar_service.update_event(
        event_id=event_id,
        subject=request.subject,
        start_datetime=request.start_datetime,
        end_datetime=request.end_datetime,
        time_zone=request.time_zone,
        body=request.body,
        body_type=request.body_type,
        location=request.location,
        attendees=request.attendees,
        is_all_day=request.is_all_day,
        is_online_meeting=request.is_online_meeting,
        reminder_minutes=request.reminder_minutes,
        show_as=request.show_as,
        importance=request.importance,
    )
    # Track which fields were updated
    changed = [k for k, v in request.model_dump().items() if v is not None]
    audit.log_calendar_update(auth.email, event_id, changed)
    return result


@router.delete("/events/{event_id}", dependencies=[Depends(require_permission("write:calendar"))])
async def delete_event(
    event_id: str,
    calendar_service: CalendarService = Depends(get_calendar_service),
    auth: Auth = Depends(get_current_auth),
):
    await calendar_service.delete_event(event_id)
    audit.log_calendar_delete(auth.email, event_id)
    return {"message": "Event deleted successfully"}


@router.post("/events/{event_id}/accept", dependencies=[Depends(require_permission("write:calendar"))])
async def accept_event(
    event_id: str,
    request: RespondEventRequest = None,
    calendar_service: CalendarService = Depends(get_calendar_service),
):
    comment = request.comment if request else None
    send_response = request.send_response if request else True
    await calendar_service.respond_to_event(
        event_id=event_id,
        response="accept",
        comment=comment,
        send_response=send_response,
    )
    return {"message": "Event accepted"}


@router.post("/events/{event_id}/tentative", dependencies=[Depends(require_permission("write:calendar"))])
async def tentatively_accept_event(
    event_id: str,
    request: RespondEventRequest = None,
    calendar_service: CalendarService = Depends(get_calendar_service),
):
    comment = request.comment if request else None
    send_response = request.send_response if request else True
    await calendar_service.respond_to_event(
        event_id=event_id,
        response="tentativelyAccept",
        comment=comment,
        send_response=send_response,
    )
    return {"message": "Event tentatively accepted"}


@router.post("/events/{event_id}/decline", dependencies=[Depends(require_permission("write:calendar"))])
async def decline_event(
    event_id: str,
    request: RespondEventRequest = None,
    calendar_service: CalendarService = Depends(get_calendar_service),
):
    comment = request.comment if request else None
    send_response = request.send_response if request else True
    await calendar_service.respond_to_event(
        event_id=event_id,
        response="decline",
        comment=comment,
        send_response=send_response,
    )
    return {"message": "Event declined"}
