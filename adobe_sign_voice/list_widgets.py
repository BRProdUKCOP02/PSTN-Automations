"""
list_widgets.py — Dev utility.

Run this script to list all Adobe Sign widgets (webforms) visible to your
integration key. It prints their IDs, names, statuses, and public URLs so
you can copy the correct ID into .env as ADOBE_SIGN_WIDGET_ID.

The widget ID appears in the webform URL as the 'wid' query parameter, e.g.:
    https://esign.adobe.com/public/esignWidget?wid=CBFCIBAA3...

Usage:
    python list_widgets.py
"""
import sys
import os
sys.path.insert(0, os.path.dirname(__file__))

from adobe_sign_client import AdobeSignClient, AdobeSignError


def main():
    print("=" * 70)
    print("Adobe Sign — Widgets (Webforms)")
    print("=" * 70)

    client = AdobeSignClient()
    try:
        widgets = client.list_widgets()
    except AdobeSignError as exc:
        print(f"\n❌  API error: {exc}")
        print("\nCheck ADOBE_SIGN_CLIENT_ID, ADOBE_SIGN_CLIENT_SECRET, "
              "ADOBE_SIGN_REFRESH_TOKEN and ADOBE_SIGN_BASE_URL in .env")
        return

    if not widgets:
        print("\nNo widgets found.")
        print("Create a webform in the Adobe Sign UI first:")
        print("  Home > Publish a web form > configure and publish")
        return

    print(f"\nFound {len(widgets)} widget(s):\n")

    for widget in widgets:
        widget_id = widget.get("id", "")
        name      = widget.get("name", "(no name)")
        status    = widget.get("status", "")

        # Try to fetch full details to get the webform URL
        url = ""
        try:
            details = client.get_widget(widget_id)
            url = details.get("url", "")
        except AdobeSignError:
            pass

        print(f"  Name   : {name}")
        print(f"  Status : {status}")
        print(f"  ID     : {widget_id}")
        if url:
            print(f"  URL    : {url}")
        print()

    print("📋  Copy the ID for your opt-out/in webform and set it as:")
    print("       ADOBE_SIGN_WIDGET_ID=<paste here>  in .env\n")


if __name__ == "__main__":
    main()
