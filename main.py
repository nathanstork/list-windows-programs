import os
import winreg


class Color:
    PURPLE = '\033[95m'
    CYAN = '\033[96m'
    DARKCYAN = '\033[36m'
    BLUE = '\033[94m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
    END = '\033[0m'


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


def get_installed_programs():
    installed_registry_programs = [
        *get_registry_programs(winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_64KEY),  # 64-bit programs
        *get_registry_programs(winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_32KEY),  # 32-bit programs
        *get_registry_programs(winreg.HKEY_CURRENT_USER, 0)  # Additional registry key for current user
    ]

    # Remove duplicates and sort programs alphabetically, ignoring case
    return sorted(list(dict.fromkeys(installed_registry_programs)), key=str.casefold)


def get_executables_from_directory(path):
    results = []

    try:
        directories = os.listdir(path)
        for directory in directories:
            # Recursively search for an executable file with the same name as the directory
            for root, dirs, files in os.walk(os.path.join(path, directory)):
                for file in files:
                    if file.endswith('.exe'):
                        # print(Color.BOLD + directory + ' > ' + Color.END + root + '\\' + file)
                        results.append(directory + ' > ' + root + '\\' + file)
                        break
    except FileNotFoundError:
        return []

    return results


def get_executables():
    program_executables = []
    # Get executables from Program Files
    program_executables += get_executables_from_directory(r'C:\Program Files')
    # Get programs from Program Files (x86)
    program_executables += get_executables_from_directory(r'C:\Program Files (x86)')
    # Get programs from ProgramData
    program_executables += get_executables_from_directory(r'C:\ProgramData')
    return program_executables


def get_environment_variables():
    environment_variables = []

    environ = os.environ
    # Format all environment variables in a readable way
    for key, value in environ.items():
        environment_variables.append(f'{key}={value}')

    return environment_variables


def write_to_txt_file(lines, filename='file'):
    # Use UTF-8 encoding to support special characters
    with open(filename + '.txt', 'w', encoding='utf-8') as file:
        for line in lines:
            file.write(line)
            if line != lines[-1]:
                file.write("\n")


if __name__ == '__main__':
    # Get installed programs
    installed_programs = get_installed_programs()
    # Print the number of installed programs in bright blue
    print(f'Found {Color.BLUE + str(len(installed_programs)) + Color.END} installed programs.')
    # Write programs to file
    write_to_txt_file(installed_programs, 'programs')

    # Get executables
    executables = get_executables()
    # Print the number of executables in bright blue
    print(f'Found {Color.YELLOW + str(len(executables)) + Color.END} additional executables.')
    # Write executables to file
    write_to_txt_file(executables, 'executables')

    # Get environment variables
    variables = get_environment_variables()
    # Print the number of environment variables in purple
    print(f'Found {Color.PURPLE + str(len(variables)) + Color.END} system variables.')
    # Write environment variables to file
    write_to_txt_file(variables, 'variables')

    print(Color.GREEN + 'Done!' + Color.END)
