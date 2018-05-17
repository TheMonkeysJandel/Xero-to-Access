# Xero-to-Access
How I managed to connect little ol MS Access to Xero using VBA

I couldn't find anyone who successfully used MS Access to connect to XERO using only VBA. And I finally nailed it.

I wish I'd recorded all my sources because I didn't do this alone. There were hundreds of people out there who's bits and pieces I took and modified to work here. Open source is amazing.

Lil note here...I'm a hobbyist programmer that wants to learn other languages but my day to day job only has me playing in VBA. My coding is ugly as shit.

This whole thing assumes the following 
- You're creating a private application for connecting XERO to A VBA source
  - You've already created your private/public key pair
- It was built in MS Access 64Bit
- You'll know what libraries/references are required to get things going (XML/Scripting etc.)
- You've got OpenSSL downloaded from http://slproweb.com/products/Win32OpenSSL.html to create your original keys. You'll need it for encoding your requests
- You're probably a real programmer

I hope this helps someone out there, it bugged me for months.
