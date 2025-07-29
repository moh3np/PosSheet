# PosSheet

This repository contains a Google Apps Script project for managing product sales.

## Debugging tips

If the product search dialog cannot find an item by code, ensure that the
inventory list is fully loaded. The dialog loads data asynchronously, so waiting
for the inventory to appear may help. If needed, the search field now falls back
to a server‑side lookup when local data is missing.

To share or back up the script, open the Apps Script editor and use **File ›
Download** to obtain a zip archive of the project.
