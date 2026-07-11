#!/usr/bin/env python3
"""
Generate a mailto: link for Outlook with proper URL encoding.
Usage: python generate_mailto.py --to "email@example.com" --subject "Subject" --body "Body text"
       python generate_mailto.py --subject "Subject" --body "Body text"  (no recipient, user fills manually)
"""

import argparse
import urllib.parse
import sys

def generate_mailto(to_addr: str, subject: str, body: str, cc: str = "", bcc: str = "") -> str:
    """
    Generate a mailto: URL with all parameters properly URL-encoded.
    The body uses \r\n for line breaks (Outlook compatible).
    """
    # Build query parameters
    params = {}
    
    if subject:
        params["subject"] = subject
    
    if body:
        # Normalize line breaks: \n → \r\n for Outlook compatibility
        normalized_body = body.replace("\r\n", "\n").replace("\n", "\r\n")
        params["body"] = normalized_body
    
    if cc:
        params["cc"] = cc
    
    if bcc:
        params["bcc"] = bcc
    
    # URL-encode query string
    # urllib.parse.urlencode uses %20 for spaces (HTML form standard)
    # For mailto:, + is also widely accepted and more compact
    query_string = urllib.parse.urlencode(params, quote_via=urllib.parse.quote)
    
    if to_addr:
        mailto_url = f"mailto:{to_addr}?{query_string}"
    else:
        mailto_url = f"mailto:?{query_string}"
    
    return mailto_url


def main():
    parser = argparse.ArgumentParser(description="Generate mailto: link for Outlook")
    parser.add_argument("--to", default="", help="Recipient email address")
    parser.add_argument("--subject", default="", help="Email subject")
    parser.add_argument("--body", default="", help="Email body text")
    parser.add_argument("--cc", default="", help="CC recipients (comma-separated)")
    parser.add_argument("--bcc", default="", help="BCC recipients (comma-separated)")
    
    args = parser.parse_args()
    
    if not args.subject and not args.body:
        # Read from stdin if no subject/body provided
        body_lines = []
        for line in sys.stdin:
            body_lines.append(line)
        body_text = "".join(body_lines).strip()
        mailto_url = generate_mailto(args.to, "", body_text, args.cc, args.bcc)
    else:
        mailto_url = generate_mailto(args.to, args.subject, args.body, args.cc, args.bcc)
    
    print(mailto_url)


if __name__ == "__main__":
    main()
