import os
import sys

try:
    import msoffcrypto
except ImportError:
    print("Error: msoffcrypto-tool is not installed.")
    print("Fix: python -m pip install msoffcrypto-tool")
    sys.exit(10)

try:
    from msoffcrypto import exceptions as msoffcrypto_exceptions
except Exception:
    msoffcrypto_exceptions = None


def is_exception_type(error, name):
    if msoffcrypto_exceptions is None:
        return False

    exception_type = getattr(msoffcrypto_exceptions, name, None)
    return exception_type is not None and isinstance(error, exception_type)


def main():
    if len(sys.argv) != 4:
        print("Error: Missing arguments.")
        print("Expected usage: python decrypt.py <input_file> <password> <output_file>")
        sys.exit(1)

    input_file = sys.argv[1]
    password = sys.argv[2]
    output_file = sys.argv[3]

    if not os.path.exists(input_file):
        print(f"Error: Input file not found: {input_file}")
        sys.exit(2)

    if password == "":
        print("Error: Excel password is empty. Please provide the correct password.")
        sys.exit(3)

    if os.path.exists(output_file):
        try:
            os.remove(output_file)
        except Exception as error:
            print(f"Error: Cannot replace existing output file: {str(error)}")
            sys.exit(4)

    try:
        with open(input_file, "rb") as file_input:
            try:
                office_file = msoffcrypto.OfficeFile(file_input)
            except Exception as error:
                print(f"Error: File cannot be opened as a Microsoft Office file: {str(error)}")
                sys.exit(5)

            try:
                if hasattr(office_file, "is_encrypted") and not office_file.is_encrypted():
                    print("Error: File is not encrypted. It should be readable without decryption.")
                    sys.exit(6)
            except Exception:
                pass

            try:
                try:
                    office_file.load_key(password=password, verify_password=True)
                except TypeError:
                    office_file.load_key(password=password)
            except Exception as error:
                message = str(error)

                if is_exception_type(error, "InvalidKeyError") or "invalid" in message.lower() or "password" in message.lower():
                    print("Error: Incorrect Excel password or unsupported encryption key.")
                    sys.exit(7)

                print(f"Error: Unable to load Excel password/key: {message}")
                sys.exit(8)

            try:
                with open(output_file, "wb") as file_output:
                    office_file.decrypt(file_output)
            except Exception as error:
                message = str(error)

                if is_exception_type(error, "DecryptionError"):
                    print(f"Error: File cannot be decrypted: {message}")
                    sys.exit(9)

                print(f"Error: Decryption failed: {message}")
                sys.exit(9)

        if not os.path.exists(output_file):
            print("Error: Decryption finished but output file was not created.")
            sys.exit(11)

        if os.path.getsize(output_file) == 0:
            print("Error: Decryption finished but output file is empty.")
            sys.exit(12)

        print("Decryption successful.")
        sys.exit(0)

    except Exception as error:
        print(f"Error: File cannot be decrypted: {str(error)}")
        sys.exit(13)


if __name__ == "__main__":
    main()
