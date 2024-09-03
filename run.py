import imaplib
import email
import re

# Konfigurasi
IMAP_SERVER = 'imap.mail.yahoo.com'
IMAP_PORT = 993
EMAIL_ACCOUNT = 'example@yahoo.com'
EMAIL_PASSWORD = 'example'

def connect_to_imap_server():
    print("Connecting to IMAP server...")
    mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
    mail.login(EMAIL_ACCOUNT, EMAIL_PASSWORD)
    print("Connected and logged in.")
    return mail

def fetch_email_headers(mail, num_headers):
    print("Selecting inbox...")
    mail.select('inbox')  # Pilih folder inbox
    print("Fetching email IDs...")
    status, data = mail.search(None, 'ALL')
    email_ids = data[0].split()
    
    # Jika jumlah email yang diminta lebih banyak dari yang tersedia, ambil semua email
    if num_headers > len(email_ids):
        num_headers = len(email_ids)
        print(f"Requested number exceeds available emails. Adjusting to {num_headers}.")
    
    headers = []

    print(f"Found {len(email_ids)} emails. Fetching the first {num_headers} headers.")
    
    for email_id in email_ids[:num_headers]:
        print(f"Fetching header for email ID: {email_id.decode()}")
        status, data = mail.fetch(email_id, '(BODY[HEADER])')
        for response_part in data:
            if isinstance(response_part, tuple):
                headers.append(response_part[1])
                
    print("Finished fetching headers.")
    return headers

def decode_header(header):
    """
    Mendekode header email yang mungkin dikodekan dengan charset tertentu.
    
    Args:
    header (str): Header email yang akan didekodekan.
    
    Returns:
    str: String yang dapat dibaca manusia.
    """
    decoded_fragments = []
    decoded_header = email.header.decode_header(header)
    
    for fragment, encoding in decoded_header:
        if isinstance(fragment, bytes):
            if encoding:
                try:
                    decoded_fragments.append(fragment.decode(encoding))
                except (TypeError, UnicodeDecodeError):
                    decoded_fragments.append(fragment.decode('utf-8', 'replace'))
            else:
                decoded_fragments.append(fragment.decode('utf-8', 'replace'))
        else:
            decoded_fragments.append(fragment)  # Already a string
    
    return ''.join(decoded_fragments)

def format_from_address(from_address):
    """
    Format alamat email dengan nama tampilan jika ada.
    
    Args:
    from_address (str): Alamat email dari header.
    
    Returns:
    str: Alamat email dalam format 'Nama <alamat-email>' jika ada nama,
         atau hanya '<alamat-email>' jika tidak ada nama.
    """
    # Gunakan regex untuk mengekstrak nama dan email
    match = re.match(r'(?P<name>.*?)\s*<(?P<email>[^>]+)>', from_address)
    if match:
        name = match.group('name').strip()
        email_address = match.group('email').strip()
        if name:
            return f'{name} <{email_address}>'
        else:
            return f'<{email_address}>'
    else:
        # Jika tidak ada nama tampilan, hanya tampilkan email
        return f'<{from_address}>'

def extract_from_email(header_data):
    msg = email.message_from_bytes(header_data)
    from_ = decode_header(msg['from'])
    return format_from_address(from_)

def save_senders_to_file(senders, filename='hasil.txt'):
    print(f"Appending senders to {filename}...")
    with open(filename, 'a', encoding='utf-8') as f:
        for sender in senders:
            f.write(f"{sender}\n")
    print("Senders appended.")

def main():
    mail = connect_to_imap_server()
    
    try:
        num_headers = int(input("Enter the number of email headers to fetch: "))
        if num_headers <= 0:
            raise ValueError("Number of headers must be a positive integer.")
    except ValueError as e:
        print(f"Invalid input: {e}")
        return
    
    email_headers = fetch_email_headers(mail, num_headers)
    
    senders = []

    print("Parsing email headers...")
    for header_data in email_headers:
        from_ = extract_from_email(header_data)
        senders.append(from_)
    
    print("Saving senders:")
    save_senders_to_file(senders)
    
    mail.logout()
    print("Logged out.")

if __name__ == "__main__":
    main()
