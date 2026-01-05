import hid

devices = hid.enumerate()
for d in devices:
    print(f"Path: {d['path']}")
    print(f"VendorID: {hex(d['vendor_id'])}, ProductID: {hex(d['product_id'])}")
    print(f"Serial: {d['serial_number']}")
    print(f"Product: {d['product_string']}")
    print(f"Manufacturer: {d['manufacturer_string']}")
    print("-" * 40)
