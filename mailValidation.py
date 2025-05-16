# Internet Is Reequired For Validating The mail ID Of The Participants.
import re
import dns.resolver

def is_valid_email(email):
    # Step 1: Basic format check
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    if not re.match(pattern, email):
        return False

    # Step 2: Domain MX record check
    domain = email.split('@')[-1]
    try:
        dns.resolver.resolve(domain, 'MX')
        return True
    except Exception:
        return False

if __name__ == "__main__":
    # Example
    print(is_valid_email("laxminayanan546@gmail.com"))    # True
    print(is_valid_email("suradalokesh2005@gmail.com"))    # True
    print(is_valid_email("nayanancsit@gmail.com"))    # True
    print(is_valid_email("nithishreddyakiti08@gmail.com"))    # True
    print(is_valid_email("test@nonexistent.tld")) # False