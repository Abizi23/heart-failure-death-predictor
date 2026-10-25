
import msal
import requests

# Replace these with the values from your Entra app's Overview page.
CLIENT_ID = "1d32c206-0938-48e7-8301-88e2a580dc56"
TENANT_ID = "8ff74247-c8d7-4028-9577-29871729fd1c"

AUTHORITY = "https://login.microsoftonline.com/common"
SCOPES = ["User.Read", "Mail.Read", "Calendars.Read"]

# Create the authentication application.
app = msal.PublicClientApplication(
    client_id=CLIENT_ID,
    authority=AUTHORITY
)

# Start Microsoft device-code sign-in.
flow = app.initiate_device_flow(scopes=SCOPES)

if "user_code" not in flow:
    raise RuntimeError(f"Could not start sign-in: {flow}")

print(flow["message"])

result = app.acquire_token_by_device_flow(flow)

if "access_token" not in result:
    raise RuntimeError(
        f"Sign-in failed: {result.get('error_description', result)}"
    )

token = result["access_token"]
headers = {"Authorization": f"Bearer {token}"}

# Test 1: Read the signed-in user's profile.
response = requests.get(
    "https://graph.microsoft.com/v1.0/me",
    headers=headers,
    timeout=30
)
response.raise_for_status()
user = response.json()

print("\nSUCCESS: Microsoft sign-in works!")
print("Signed-in account:", user.get("userPrincipalName"))

# Test 2: Read up to five recent emails.
response = requests.get(
    "https://graph.microsoft.com/v1.0/me/messages",
    headers=headers,
    params={
        "$select": "subject,receivedDateTime,bodyPreview",
        "$top": 5,
        "$orderby": "receivedDateTime desc"
    },
    timeout=30
)
if not response.ok:
    print("Email request failed:", response.status_code)
    print("Details:", response.text)
    raise SystemExit("Could not read Outlook emails.")

print("\nRECENT EMAILS")
for email in response.json().get("value", []):
    print("-", email.get("subject", "(No subject)"))
    print("  Received:", email.get("receivedDateTime"))

# Test 3: Read up to five upcoming calendar events.
response = requests.get(
    "https://graph.microsoft.com/v1.0/me/events",
    headers=headers,
    params={
        "$select": "subject,start,end",
        "$top": 5,
        "$orderby": "start/dateTime"
    },
    timeout=30
)
response.raise_for_status()

print("\nCALENDAR EVENTS")
for event in response.json().get("value", []):
    print("-", event.get("subject", "(No subject)"))
    print("  Start:", event.get("start", {}).get("dateTime"))

print("\nSUCCESS: Outlook email and calendar access tested.")
