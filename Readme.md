# KMTronic 4-Channel USB Relay Control

## Description

**Relay_Control.exe** controls a 4-Channel USB relay module.
The UI allows you to turn **ON** or **OFF** the relay channels individually or all together.

---

## Requirements

1. Install the following Python packages before running the script:

   ```bash
   pip install pyserial openpyxl

Note: tkinter and os are part of the standard Python library.

Install VCP Drivers from: <https://ftdichip.com/drivers/vcp-drivers/>

## Usage

1. Run Relay_Control.exe.
2. Select corresponding COM port the KMTronic relay module using the pull
 down menu.
3. Click on 'Turn ON RLx' and 'Turn OFF RLx' buttons to control the
 corresponding relay channels
4. The Status circles reflects the relay status
  RED  - Relay OFF
  GREEN - Relay ON
5. Set All button turn ON all the relay channels
6. Reset All button turn OFF all relay channels
7. If the relay module is disconnected, the relay and the status are
 resetted. Perform step2 before starting operating the UI.

## Troubleshooting

1. Status and actual relay states are different.
   1. Wrong USB port might have selected.
      a.Identify correct COM port from device manager and select it from the pull down menu.
   2. KMtronic relay module might have disconnected while the exe is not   running.
      b. Select proper COM port from the pull down menu, Reset All and proceed using the UI.
   3. relay_log.xlsx might have got deleted
      c. Launch Relay_Control.exe, select proper COM port, Reset All and proceed using the UI.
