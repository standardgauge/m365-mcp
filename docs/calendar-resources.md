# Calendar resources: rooms and equipment

How the calendar tools book a room or equipment mailbox, how a caller finds
one, and the part Microsoft Graph does not cover.

## Booking a resource

Exchange has two kinds of resource mailbox: **room** mailboxes (meeting rooms)
and **equipment** mailboxes (a pool car, a projector, a camera kit). Both are
booked the same way, by inviting the mailbox to the event as an attendee of
type `resource`.

`create_event` and `update_event` take each attendee in one of two forms:

```json
{
  "attendees": [
    "alex@contoso.com",
    { "email": "pat@contoso.com", "type": "optional" },
    { "email": "poolcar@contoso.com", "type": "resource" }
  ]
}
```

A bare address is a `required` attendee, which is what every attendee was
before the typed form existed, so existing callers are unchanged. The object
form takes `email` and an optional `type` of `required`, `optional` or
`resource`. Anything else is refused before Graph is called: an unknown type,
an unknown field, a missing address, or the same address twice.

The type matters. A resource invited as `required` gets an ordinary meeting
request, which nothing processes, so the booking never happens. Invited as
`resource`, the mailbox's booking policy runs: depending on how the mailbox is
configured it accepts, declines a conflict, or holds the request for a
delegate.

On `update_event`, `attendees` replaces the whole list, as it always has.

## Seeing whether it accepted

`get_event` returns each attendee as an object rather than a bare address:

```json
{
  "email": "poolcar@contoso.com",
  "name": "Pool car",
  "type": "resource",
  "response": "accepted",
  "respondedAt": "2026-07-01T09:00:05Z"
}
```

`response` is Graph's value: `none`, `organizer`, `tentativelyAccepted`,
`accepted`, `declined` or `notResponded`. A resource usually answers within
seconds of the invitation, so read the event again shortly after booking. A
`declined` resource is not booked; pick another time or another resource.
`respondedAt` is null until the attendee has responded.

`list_events` still returns bare attendee addresses, to keep the listing
small. Use `get_event` for the detail.

## Finding a room

`list_rooms` reads `GET /places/microsoft.graph.room`. That needs the
delegated permission `Place.Read.All` with admin consent on the app
registration. The README's permission table lists it, and
[Entra setup](entra-setup.md#verifying-consent-actually-landed) shows how to
check the consent actually landed. Without it Graph answers 403, and `list_rooms` says which permission
is missing instead of passing on Graph's generic "Access is denied". Booking a
room whose address you already know does not need `Place.Read.All`; only the
listing does.

## Finding equipment: what Graph cannot do

There is no Graph call that lists equipment mailboxes.

- The places API (`/places`) covers rooms, room lists and workspaces. An
  equipment mailbox is not a place, so it never appears there, however it is
  configured.
- The user object Graph returns for a resource mailbox does not say what kind
  of mailbox it is. Exchange's recipient type (`EquipmentMailbox`) is not
  exposed through the delegated Graph surface this server uses.
- `mailboxSettings.userPurpose` does report `equipment`, but reading another
  mailbox's settings needs access to that mailbox, which a delegated caller
  does not have.

So a caller has to know an equipment mailbox's address to book it. In
practice:

- An administrator lists them in Exchange Online PowerShell with
  `Get-Mailbox -RecipientTypeDetails EquipmentMailbox`, or in the Exchange
  admin center under **Resources**, and shares the addresses with the people
  who book them.
- The address can be saved as a contact, so `search_contacts` finds it by
  name.
- Once the address is known, `get_schedule` reads the resource's free/busy
  like anyone else's, so a caller can check it is free before booking.

If Microsoft adds equipment to the places API, `list_rooms` is the place to
extend.
