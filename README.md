# ASCOM.NGCAT.Focuser
Robofocus controller through ASCOM

Having a Robofocus, I ran into a few problems with the various drivers I could find:
- v1 drivers didn't have access to temperature.
- v3 drivers have temperature but don't save settings well and are a bit unstable.

Lack of development from the manufacturer for many years means that the focuser is abandonned. So I found an old documentation that mentions the protocol and implemented my own.

It is very simple. It doesn't allow you to set the options of the focuser (but I don't see why you would need to). However, the code could do it. It just needs a UI for that part.

No support, no guarantees. This project is only here to help those who would want to build their own driver.