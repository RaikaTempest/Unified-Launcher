***INSTALL***



Copy the whole folder somewhere local



To use this you will require python. Open the Microsoft store and search python and download. Just grab the latest, it won't require IT to allow.



Once installed, press "WIN+R" to open the run command and type CMD and press enter. This will open command prompt.

We need to install some dependencies, copy and paste each of these into command prompt and press enter, one at a time.



pip install tqdm

pip install pandas

pip install openpyxl

pip install Pillow pillow-heif



*(You may get some warnings about write access, don't worry this application does not require any privileges so we can ignore those.)*



Now you can double click UPR.py to run the application.





***USE***



Click file -> new review session



In the first window, select the field inspection folder with the photos

In the second window, select the CSV with the barcodes provided by the ISP (in tender documents) for the whole PAP. (I like to copy it to the photos folder before I begin)



Now wait. This will take a minute. You can view the command prompt window to see progress, but it will scan the DR and PN list for matches to the barcode table

Following this it will copy all the photos locally for easier review. Once complete you can begin.



Note: You can click and drag to re-order the barcodes on the left (it's a bit laggy, sorry), this will determine the order of the report.

