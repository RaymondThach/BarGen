A python application to scrape reference numbers of products from PDF planograms of special ends for supermarkets. Scrapped reference numbers are converted into barcodes to efficiently scan rather than type in each time. The planogram must refer to the reference numbers in the format of "Ref: NUM" or "Ref: NUM1, NUM2, NUM3", where NUM is the reference numbers. This is designed as a quick temporary solution not a permanent one.

Instructions: 
1. Place at the root of the file the PDF of the planogram named "plan.pdf"
2. Run the bargen.py code, and it should extract the references numbers and generate a sorted Word Document of the bar codes
3. Cross check the barcodes and reference numbers. If any are missing readjust the PDF regions as some PDF's may be of different resolutions.
4. Edit the final Word Document as needed.