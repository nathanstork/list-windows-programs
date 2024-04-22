import winreg


def get_registry_programs(hive, flag):
    programs = []

    # Open the registry key containing information about installed programs
    key = winreg.OpenKey(hive, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", flag)

    try:
        index = 0
        while True:
            # Enumerate subkeys to access information about programs
            subkey_name = winreg.EnumKey(key, index)
            subkey = winreg.OpenKey(key, subkey_name)

            try:
                # Read the DisplayName value to get the program name
                program_name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                programs.append(program_name)
            except FileNotFoundError:
                pass  # Some subkeys might not have the DisplayName value

            index += 1
    except OSError as e:
        if e.errno == 22:  # Error code for no more data is available
            pass
        else:
            print(f"An error occurred: {e}")
    finally:
        winreg.CloseKey(key)

    return programs


def get_programs_from_program_files():
    # List all folders in the Program Files directory

    pass


def get_installed_programs():
    installed_registry_programs = [
        *get_registry_programs(winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_64KEY),  # 64-bit programs
        *get_registry_programs(winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_32KEY),  # 32-bit programs
        *get_registry_programs(winreg.HKEY_CURRENT_USER, 0)  # Additional registry key for current user
    ]

    # Remove duplicates and sort programs alphabetically, ignoring case
    return sorted(list(dict.fromkeys(installed_registry_programs)), key=str.casefold)


def write_to_file(programs):
    # Use UTF-8 encoding to support special characters
    with open("installed_programs.txt", "w", encoding="utf-8") as file:
        for p in programs:
            file.write(p)
            if p != programs[-1]:
                file.write("\n")


if __name__ == '__main__':
    # Get installed programs
    installed_programs = get_installed_programs()
    # Print the number of installed programs in bright blue
    print(f"Found \033[94m{len(installed_programs)}\033[0m installed programs.")
    # FOR TESTING PURPOSES: Print the list of installed programs
    for program in installed_programs:
        print(program)
    # Write to file
    write_to_file(installed_programs)
