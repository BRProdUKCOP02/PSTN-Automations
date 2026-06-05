"""
list_library_docs.py — Dev utility.

Run this script once after building your template in the Adobe Sign UI.
It lists all Library Documents visible to your integration key and prints
their IDs and names so you can copy the correct ID into .env.

Usage:
    python list_library_docs.py
"""
import sys
import os
sys.path.insert(0, os.path.dirname(__file__))

from adobe_sign_client import AdobeSignClient, AdobeSignError


def main():
    print("=" * 60)
    print("Adobe Sign — Library Documents")
    print("=" * 60)

    client = AdobeSignClient()
    try:
        docs = client.list_library_documents()
    except AdobeSignError as exc:
        print(f"\n❌  API error: {exc}")
        print("\nCheck ADOBE_SIGN_INTEGRATION_KEY and ADOBE_SIGN_BASE_URL in .env")
        return

    if not docs:
        print("\nNo library documents found.")
        print("Build your template in the Adobe Sign UI first:")
        print("  Home > Start from library > Create a template")
        return

    print(f"\nFound {len(docs)} library document(s):\n")
    print(f"  {'ID':<50}  {'Name'}")
    print(f"  {'-'*50}  {'-'*40}")
    for doc in docs:
        doc_id = doc.get("id", "")
        name = doc.get("name", "")
        print(f"  {doc_id:<50}  {name}")

    print("\n📋  Copy the ID for your opt-out/in template and set it as:")
    print("       ADOBE_SIGN_LIBRARY_DOC_ID=<paste here>  in .env\n")


if __name__ == "__main__":
    main()
