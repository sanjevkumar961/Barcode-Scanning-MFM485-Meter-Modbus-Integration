#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HID Device Enumeration Test Utility

Author: Sanjev Kumar (SK96)
License: Apache License 2.0

Copyright 2026 Sanjev Kumar

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
"""

import hid

devices = hid.enumerate()
for d in devices:
    print(f"Path: {d['path']}")
    print(f"VendorID: {hex(d['vendor_id'])}, ProductID: {hex(d['product_id'])}")
    print(f"Serial: {d['serial_number']}")
    print(f"Product: {d['product_string']}")
    print(f"Manufacturer: {d['manufacturer_string']}")
    print("-" * 40)
