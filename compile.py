import subprocess
import os
import sys

try:
    from termcolor import colored
    import colorama

    colorama.init()
except Exception:
    colors_available = False
else:
    colors_available = True


def print_error(message):
    if colors_available:
        print(colored(message, 'red'))
    else:
        print(message)


def ask_yes_no(question):
    user_input = ''
    while user_input.lower() not in ['y', 'n']:
        user_input = input(question + ' (y/n)? ')
    if user_input.lower() == 'y':
        return True
    else:
        return False


def main():
    print('#########################################')
    print('#                                       #')
    print('#  Software Updater compilation script  #')
    print('#                                       #')
    print('#########################################\n')

    print('Preparing...', end='\r')

    with open('requirements.txt') as file:
        requirements = ' '.join(file.read().split('\n') + ['pyinstaller'])

    s = subprocess.run(f'pip show {requirements}',
                       stdout=subprocess.PIPE,
                       stderr=subprocess.STDOUT,
                       creationflags=subprocess.CREATE_NO_WINDOW)
    output = s.stdout.decode()

    if 'WARNING: Package(s) not found' in output:
        requirements_installed = False
    else:
        requirements_installed = True

    try:
        subprocess.run('makensis',
                       stdout=subprocess.PIPE,
                       stderr=subprocess.STDOUT,
                       creationflags=subprocess.CREATE_NO_WINDOW)
    except FileNotFoundError:
        nsis_available = False
    else:
        nsis_available = True

    if not requirements_installed:
        print_error('You do not have installed the required Python packages.')

        if ask_yes_no('Do you want to install them'):
            os.system(f'pip install {requirements}')

            os.system(f'py {__file__}')

        sys.exit(0)

    if not nsis_available:
        print_error(
            'You do not have installed NSIS or you do not have added its location to PATH environment variable')

    if nsis_available:
        create_installer = ask_yes_no(
            'Do you want to create installer after compilation?')
    else:
        create_installer = False

    print('\nCompiling...')
    if os.path.isfile(os.path.join(os.path.split(__file__)[0],
                                   'SoftwareUpdater.spec')):
        print('Using existing spec file.')

        os.system('pyinstaller SoftwareUpdater.spec')
    else:
        print('PyInstaller will create spec file because it does not exist')

        s = subprocess.run('pip show customtkinter',
                           stdout=subprocess.PIPE,
                           stderr=subprocess.STDOUT,
                           creationflags=subprocess.CREATE_NO_WINDOW)
        output = s.stdout.decode().split('\r\n')

        for line in output:
            if line.startswith('Location: '):
                customtkinter_path = os.path.join(
                    line.split('Location: ')[1], 'customtkinter')

        os.system(
            f'pyinstaller --noconsole --onefile --windowed --icon="img/elephant.ico" --hidden-import=customtkinter --hidden-import=PIL --hidden-import=requests --add-data="img;img/" --add-data="{customtkinter_path};customtkinter/" SoftwareUpdater.py')

    if create_installer:
        print('\nCreating installer...')

        os.system('makensis setup.nsi')

    if create_installer:
        print('\nCompiled file and installer should be in "dist" folder.')
    else:
        print('\nCompiled file should be in "dist" folder.')


if __name__ == '__main__':
    main()
