''' TUI for OverBytes Toolbox '''
import subprocess
from textual import on
from textual.app import App
from textual.containers import Container, VerticalScroll
from textual.widgets import Button, Footer, Header, Static
from ascii import ChonkCatAscii, ElGatoAscii, NerdCatAscii


class AsciiTitle(Static):
    '''The title in ascii'''

    def compose(self):
        yield Static("""
________                   __________          __                     .____    .____   _________  
\\_____  \\___  __ __________\\______   \\___.__._/  |_  ____   ______    |    |   |    |  \\_   ___ \\
 /   |   \\  \\/ // __ \\_  __ \\    |  _<   |  |\\   __\\/ __ \\ /  ___/    |    |   |    |  /    \\  \\/
/    |    \\   /\\  ___/|  | \\/    |   \\\\___  | |  | \\  ___/ \\___ \\     |    |___|    |__\\     \\____
\\_______  /\\_/  \\___  |__|  |______  // ____| |__|  \\___  >____  > /\\ |_______ \\_______ \\______  /
        \\/          \\/             \\/ \\/                \\/     \\/  \\/         \\/       \\/      \\/
""")


class ScriptButton(Static):
    '''A class (like a template) for the script buttons'''

    def __init__(self, powershell_command: str, button_name: str, hidden=False, **kwargs):
        super().__init__(**kwargs)
        self.button_name = button_name
        self.powershell_command = powershell_command
        self.hidden = hidden

    @on(Button.Pressed)
    def set_command(self) -> None:
        """When this button is pressed, tell the app which command to use."""
        if not self.hidden:
            self.app.current_command = f'Start-Process powershell.exe -ArgumentList "-NoExit", "-Command", \
                "Import-Module \'$env:TEMP\\obsoftware\\ob.psm1\'; Import-Module \'$env:TEMP\\obsoftware\\dentalsoftware.psm1\';\
                    {self.powershell_command}; Start-Sleep -Seconds 3; Exit"'
        else:
            # The hidden functionality is not yet functional. If I knew why, it would be functional
            self.app.current_command = f"Start-Job -ScriptBlock {{ Import-Module \"$env:TEMP\\obsoftware\\ob.psm1\"; {self.powershell_command}; Exit }} > $null"
        self.app.current_ps_command = f"{self.powershell_command}"
        self.app.sub_title = f"Selected: {self.button_name}"

    def compose(self):
        yield Button(f"{self.button_name}")


class ESIOSSButton(Static):
    """Button for going to the ES IOSS Configuration Section"""
    @on(Button.Pressed)
    def ioss_config_section(self) -> None:
        """Switches to the ES IOSS Config Section"""
        self.app.switch_section(
            "#es-ioss-buttons", "Eaglesoft IOSS Config Section")

    def compose(self):
        yield Button("ES IOSS Configuration")


class OverBytesToolbox(App):
    '''TUI for obtoolbox'''
    BINDINGS = [
        # (key, action_name, description),
        # omit the action_ prefix for the method
        ("d", "toggle_dark", "Toggle Theme"),
        ("x", "execute_command", "Execute Command"),
        ("h", "main_section", "Home"),
        ("s", "dental_software_section", "Dental Software"),
        ("q", "exit_obtoolbox", "Exit")
    ]
    CSS_PATH = "obtoolbox.tcss"

    def __init__(self, driver_class=None, css_path=None, watch_css=False, ansi_color=False):
        super().__init__(driver_class, css_path, watch_css, ansi_color)
        self.current_command = None
        self.current_ps_command = None
        self.sub_title = ""

    def compose(self):
        yield Header()
        yield Footer()
        yield AsciiTitle()
        with Container(id="main-buttons"):
            with VerticalScroll(id="left-column"):
                yield ScriptButton(button_name="Release & Renew IP",
                                   powershell_command="Invoke-ReleaseRenew")
                yield ScriptButton(button_name="Windows Repair",
                                   powershell_command="Invoke-WindowsRepair")
                yield ScriptButton(button_name="Reset Print Spooler",
                                   powershell_command="Invoke-ResetSpooler")
                yield ScriptButton(button_name="Get Serial Number",
                                   powershell_command="Get-DeviceInfo")
                yield ScriptButton(button_name="Run GPUpdate",
                                   powershell_command="Update-GroupPolicy")
                yield ScriptButton(button_name="Flush and Register DNS",
                                   powershell_command="Invoke-ReleaseRenew")
                yield ScriptButton(button_name="Create Local Admin User",
                                   powershell_command="New-LocalAdmin")
                yield ScriptButton(button_name="Enable PSRemoting",
                                   powershell_command="Invoke-PSRemoting")
                yield ScriptButton(button_name="Disable IPv6",
                                   powershell_command="Disable-IPv6")
                yield ScriptButton(button_name="Enable Remote Desktop",
                                   powershell_command="Enable-RDP")
                yield ScriptButton(button_name="Firewall | Allow All",
                                   powershell_command="Unlock-FirewallAll")
                yield ScriptButton(button_name="ChkDsk C:\\",
                                   powershell_command="Invoke-ChkdskC")
                yield ScriptButton(button_name="Install C++ 2015-2022",
                                   powershell_command="Get-CppRedist")
            yield NerdCatAscii(classes="ascii-art")
            with VerticalScroll(id="right-column"):
                yield ScriptButton(button_name="Run CLI Net Scan",
                                   powershell_command="Invoke-NetScan")
                yield ScriptButton(button_name="Advanced IP Scanner",
                                   powershell_command="Get-IPScanner")
                yield ScriptButton(button_name="WizTree",
                                   powershell_command="Invoke-WizTree")
                yield ScriptButton(button_name=".Net Repair Tool",
                                   powershell_command="Invoke-DotNetRepair")
                yield ScriptButton(button_name="HWInfo",
                                   powershell_command="Invoke-HWInfo")
                yield ScriptButton(button_name="TeamViewerQS",
                                   powershell_command="Invoke-TeamViewerQS")
                yield ScriptButton(button_name="Revo Uninstaller",
                                   powershell_command="Invoke-RevoUninstaller")
                yield ScriptButton(button_name="Office Install",
                                   powershell_command="Invoke-OfficeInstall")
                yield ScriptButton(button_name="Hold Music",
                                   powershell_command="Invoke-HoldMusicSection")
        with Container(id="dental-buttons",
                       classes="hidden"):
            with VerticalScroll(id="left-column"):
                yield Static("Patterson Dental")
                yield ScriptButton(button_name="Eaglesoft General Fix",
                                   powershell_command="Stop-EaglesoftProcesses; Invoke-ESServConnectFix; Invoke-ESHexDecFix; Invoke-OCXREGFix; Invoke-Eaglesoft")
                yield ScriptButton(button_name="Reinstall SmartDoc Printer",
                                   powershell_command="Invoke-SmartDocScannerFix")
                yield ScriptButton(button_name="Eaglesoft Download Page",
                                   powershell_command="Open-ExternalLink 'https://pattersonsupport.custhelp.com/app/answers/detail/a_id/23400#New%20Server'")
                yield ScriptButton(button_name="Schick with ES24.20+ Page",
                                   powershell_command="Open-ExternalLink 'https://pattersonsupport.custhelp.com/app/answers/detail/a_id/44313/kw/44313'")
                yield ScriptButton(button_name="Eaglesoft Dexis/Gendex sensor integration",
                                   powershell_command="Install-CDREliteDriver")
                yield ESIOSSButton()
                yield Static("VATECH")
                yield ScriptButton(button_name="libiomp5md.dll error fix",
                                   powershell_command="Invoke-EZDentDLLFix")
                yield ScriptButton(button_name="Clear EzDent Cache",
                                   powershell_command="Clear-EZCache")
            yield ElGatoAscii(classes="ascii-art")
            with VerticalScroll(id="right-column"):
                yield Static("DentSply Sirona")
                yield ScriptButton(button_name="Schick Drivers Page",
                                   powershell_command="Open-ExternalLink 'https://www.dentsplysironasupport.com/en-us/user_section/user_section_imaging/schick_brand_software.html'")
                yield ScriptButton(button_name="Sidexis Migration Section for 4.3",
                                   powershell_command="Open-SidexisMigrationSection")
                yield Static("Others")
                yield ScriptButton(button_name="Install Mouthwatch Drivers",
                                   powershell_command="Install-MouthwatchDrivers")
                yield ScriptButton(button_name="TDO XDR Sensor Fix for Win11",
                                   powershell_command="Invoke-TDOXDRW11Fix")
        with Container(id="es-ioss-buttons",
                       classes="hidden"):
            with VerticalScroll(id="left-column"):
                yield ScriptButton(button_name="1. Disable Memory Integrity",
                                   powershell_command="Disable-MemoryIntegrity")
                yield ScriptButton(button_name="2. MSXML 4.0 Install",
                                   powershell_command="Install-MSXML4")
                yield ScriptButton(button_name="3. Install IOSS",
                                   powershell_command="Install-IOSS")
                yield ScriptButton(button_name="4. Optimize the IOSS Service",
                                   powershell_command="Optimize-IOSSService")
                yield ScriptButton(button_name="5. Install CDRElite Driver",
                                   powershell_command="Install-CDREliteDriver")
                yield ScriptButton(button_name="6. CDRElite Patching and Configuring ES24.20+",
                                   powershell_command="Install-CDRPatch")
                yield ScriptButton(button_name="7. Schick Sensor Integration for OPs",
                                   powershell_command="Install-ESSchickSensorIntegration")
            yield ChonkCatAscii(classes="ascii-art")

    def on_mount(self):
        ''' Things to run in preperation of the script '''

    def action_execute_command(self) -> None:
        """Runs the currently selected command when 'x' is pressed."""
        if not self.current_command:
            self.sub_title = "No command selected!"
            return
        subprocess.run(["powershell", "-Command",
                       self.current_command], check=False)
        self.sub_title = f"Executed: {self.current_ps_command}"

    def action_dental_software_section(self) -> None:
        """Switches to the dental button grid"""
        self.switch_section("#dental-buttons", "Dental Software Section")

    def action_main_section(self) -> None:
        """Switches to the main section"""
        self.switch_section("#main-buttons", "Home")

    def action_exit_obtoolbox(self) -> None:
        """Exits the application cleanly"""
        self.exit()

    def switch_section(self, section_id: str, subtitle: str) -> None:
        """Hides all button sets but the specified section_id"""
        for button_set in ["#dental-buttons", "#main-buttons", "#es-ioss-buttons"]:
            grid = self.query_one(button_set)
            if button_set == section_id:
                grid.remove_class("hidden")
            else:
                grid.add_class("hidden")
        self.sub_title = subtitle


def main():
    """The main executive when importing into other python files"""
    OverBytesToolbox().run()


if __name__ == '__main__':
    main()
