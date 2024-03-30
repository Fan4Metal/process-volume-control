import sys
import keyboard
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
import pythoncom

KEY_VOLUME_DOWN = -174
KEY_VOLUME_UP = -175


def get_process_step():
    args = sys.argv
    if len(args) != 3:
        print("Usage example:\n\t python proc_vol.py \"winamp.exe\" 0.025")
        exit(1)
    try:
        process, step = str(args[1]), float(args[2])
    except ValueError as e:
        print(e)
        exit(1)
    return process, step


def get_volume_obj(process):
    pythoncom.CoInitialize()
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        if session.Process and session.Process.name() == process:
            volume_obj = session._ctl.QueryInterface(ISimpleAudioVolume)
            return volume_obj


def volume_up(process, step):
    volume_obj = get_volume_obj(process)
    if not volume_obj:
        return
    current_volume = volume_obj.GetMasterVolume()
    new_volume = current_volume + step
    if new_volume > 1.0:
        new_volume = 1.0
    volume_obj.SetMasterVolume(new_volume, None)


def volume_down(process, step):
    volume_obj = get_volume_obj(process)
    if not volume_obj:
        return
    current_volume = volume_obj.GetMasterVolume()
    new_volume = current_volume - step
    if new_volume < 0:
        new_volume = 0
    volume_obj.SetMasterVolume(new_volume, None)


def main():
    process, step = get_process_step()
    keyboard.add_hotkey(KEY_VOLUME_DOWN, volume_down, args=[process, step], suppress=True)
    keyboard.add_hotkey(KEY_VOLUME_UP, volume_up, args=[process, step], suppress=True)
    keyboard.wait()


if __name__ == "__main__":
    main()
