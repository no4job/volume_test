"""
Get and set access to master volume example.
"""

from __future__ import print_function
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume,ISimpleAudioVolume


def main():
    devices = AudioUtilities.GetSpeakers()
    all_devices = AudioUtilities.GetAllDevices()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    # print(volume)
    print("volume.QueryHardwareSupport(): %s" % volume.QueryHardwareSupport())
    print("volume.GetMasterVolumeLevelScalar(): %s" % volume.GetMasterVolumeLevelScalar())
    print("volume.GetMasterVolumeLevel(): %s" % volume.GetMasterVolumeLevel())
    print(volume.GetMasterVolumeLevel())
    print("volume.GetVolumeRange(): (%s, %s, %s)" % volume.GetVolumeRange())
    print("volume.GetVolumeStepInfo(): (%s,%s)" % volume.GetVolumeStepInfo())

    # interface2 = devices.Activate(
    #     IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        # volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        volume = session.SimpleAudioVolume
        if session.Process:
            print(session.Process.name())
        if session.Process and session.Process.name() == "python.exe":
            print("fprocess python.exe volume.GetMasterVolume(): %s" % volume.GetMasterVolume())


'''
https://gist.github.com/v3nturetheworld/b43c775a4fa52a8bd8cef0d50a6bc79a
'''
if __name__ == "__main__":
    main()