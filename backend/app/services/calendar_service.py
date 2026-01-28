from typing import Optional
from datetime import datetime
from app.services.graph_client import GraphClient


class CalendarService:
    def __init__(self, graph_client: GraphClient):
        self.client = graph_client

    async def list_calendars(self) -> dict:
        return await self.client.get("/me/calendars")

    async def list_events(
        self,
        calendar_id: Optional[str] = None,
        top: int = 25,
        skip: int = 0,
        order_by: str = "start/dateTime",
        filter_query: Optional[str] = None,
    ) -> dict:
        if calendar_id:
            endpoint = f"/me/calendars/{calendar_id}/events"
        else:
            endpoint = "/me/events"

        params = {
            "$top": top,
            "$skip": skip,
            "$orderby": order_by,
        }

        if filter_query:
            params["$filter"] = filter_query

        return await self.client.get(endpoint, params=params)

    async def get_calendar_view(
        self,
        start_datetime: datetime,
        end_datetime: datetime,
        calendar_id: Optional[str] = None,
        top: int = 100,
    ) -> dict:
        if calendar_id:
            endpoint = f"/me/calendars/{calendar_id}/calendarView"
        else:
            endpoint = "/me/calendarView"

        params = {
            "startDateTime": start_datetime.isoformat() + "Z",
            "endDateTime": end_datetime.isoformat() + "Z",
            "$top": top,
        }

        return await self.client.get(endpoint, params=params)

    async def get_event(self, event_id: str) -> dict:
        return await self.client.get(f"/me/events/{event_id}")

    async def create_event(
        self,
        subject: str,
        start_datetime: datetime,
        end_datetime: datetime,
        time_zone: str = "UTC",
        body: Optional[str] = None,
        body_type: str = "HTML",
        location: Optional[str] = None,
        attendees: list[str] = [],
        is_all_day: bool = False,
        is_online_meeting: bool = False,
        recurrence: Optional[dict] = None,
        reminder_minutes: int = 15,
        show_as: str = "busy",
        importance: str = "normal",
        calendar_id: Optional[str] = None,
    ) -> dict:
        event_data = {
            "subject": subject,
            "start": {
                "dateTime": start_datetime.isoformat(),
                "timeZone": time_zone,
            },
            "end": {
                "dateTime": end_datetime.isoformat(),
                "timeZone": time_zone,
            },
            "isAllDay": is_all_day,
            "isOnlineMeeting": is_online_meeting,
            "reminderMinutesBeforeStart": reminder_minutes,
            "showAs": show_as,
            "importance": importance,
        }

        if body:
            event_data["body"] = {
                "contentType": body_type,
                "content": body,
            }

        if location:
            event_data["location"] = {"displayName": location}

        if attendees:
            event_data["attendees"] = [
                {
                    "emailAddress": {"address": email},
                    "type": "required",
                }
                for email in attendees
            ]

        if recurrence:
            event_data["recurrence"] = recurrence

        if calendar_id:
            endpoint = f"/me/calendars/{calendar_id}/events"
        else:
            endpoint = "/me/events"

        return await self.client.post(endpoint, event_data)

    async def update_event(
        self,
        event_id: str,
        subject: Optional[str] = None,
        start_datetime: Optional[datetime] = None,
        end_datetime: Optional[datetime] = None,
        time_zone: str = "UTC",
        body: Optional[str] = None,
        body_type: str = "HTML",
        location: Optional[str] = None,
        attendees: Optional[list[str]] = None,
        is_all_day: Optional[bool] = None,
        is_online_meeting: Optional[bool] = None,
        reminder_minutes: Optional[int] = None,
        show_as: Optional[str] = None,
        importance: Optional[str] = None,
    ) -> dict:
        event_data = {}

        if subject is not None:
            event_data["subject"] = subject

        if start_datetime is not None:
            event_data["start"] = {
                "dateTime": start_datetime.isoformat(),
                "timeZone": time_zone,
            }

        if end_datetime is not None:
            event_data["end"] = {
                "dateTime": end_datetime.isoformat(),
                "timeZone": time_zone,
            }

        if body is not None:
            event_data["body"] = {
                "contentType": body_type,
                "content": body,
            }

        if location is not None:
            event_data["location"] = {"displayName": location}

        if attendees is not None:
            event_data["attendees"] = [
                {
                    "emailAddress": {"address": email},
                    "type": "required",
                }
                for email in attendees
            ]

        if is_all_day is not None:
            event_data["isAllDay"] = is_all_day

        if is_online_meeting is not None:
            event_data["isOnlineMeeting"] = is_online_meeting

        if reminder_minutes is not None:
            event_data["reminderMinutesBeforeStart"] = reminder_minutes

        if show_as is not None:
            event_data["showAs"] = show_as

        if importance is not None:
            event_data["importance"] = importance

        return await self.client.patch(f"/me/events/{event_id}", event_data)

    async def delete_event(self, event_id: str) -> None:
        await self.client.delete(f"/me/events/{event_id}")

    async def respond_to_event(
        self,
        event_id: str,
        response: str,  # accept, tentativelyAccept, decline
        comment: Optional[str] = None,
        send_response: bool = True,
    ) -> dict:
        data = {"sendResponse": send_response}
        if comment:
            data["comment"] = comment

        return await self.client.post(f"/me/events/{event_id}/{response}", data)
