import serial

class Device:    
    @property
    def type(self):
        return self._type

    @type.setter
    def type(self, type):
        self._type = type

class DeviceBarcodeScanner(Device):
    @property
    def mode(self):
        return self._mode

    @mode.setter
    def mode(self, mode):
        self._mode = mode

    @property
    def port(self):
        return self._port

    @port.setter
    def port(self, port):
        self._port = port

    @property
    def baudrate(self):
        return self._baudrate

    @baudrate.setter
    def baudrate(self, baudrate):
        self._baudrate = baudrate

    @property
    def suffix(self):
        return self._suffix

    @suffix.setter
    def suffix(self, suffix):
        self._suffix = suffix

def get_raw_data(device:DeviceBarcodeScanner):
    scanner = serial.Serial(port=device.port, baudrate=device.baudrate)
    raw_line = scanner.read_until(expected=device.suffix)
    if raw_line:
        text = raw_line.decode("utf-8", errors="ignore").strip()
        return text

def get_ki(device:DeviceBarcodeScanner):
    raw_data = get_raw_data(device)
    km_prefix = raw_data[0:2]
    gtin = raw_data[2:16]
    serial_index_start = 18
    serial_index_end = raw_data[serial_index_start:].find("\x1d")+serial_index_start
    serial = raw_data[serial_index_start:serial_index_end]
    ki = km_prefix + gtin + '21' + serial
    return ki