
Script name: Imprinted-gene-catcher

Description and tasks: This script is written to work with loss of heterozygosity (LOH) data from Chromosome Analysis Suite 4.3 (ChAS 4.3) 

Working in the program (ChAS) we analyze the loss of heterozygosity (LOH).
We have 2 tasks: Determination of loss of heterozygosity LOH above 2500000 (we will assume that the size above this is abnormal) bp and determination of imprinted genes (gene that is differentially expressed depending on maternal or paternal origin ). 

Important point.

The LOH size that is considered critical can be changed, and the list of genes that are searched can also be changed.


1-I wrote the code in PyCharm. Need to install #install python-docx.
This package is needed to generate a doc file and highlight text with color.

2-The geneimprint section provides a list of imprinted genes. I took this list from www.geneimprint.com. The list of genes can be changed to your taste.
 
3-First you enter how much LOH there is in the your program. For example :1 

4-Next, you will copy all the data for this LOH. Columns from the CHAS program should be copied 14.
Below are the names of these cells and an example row:

1:PatientName 2:Type 3:Chromosome 4:Min 5:Max 6:Size 7:M.Count 8:M.M.Dis 9:Cyto.Start 10:Cyto.End 11:Gene.Count 12:Omim.G.Count 13:Genes 14:Omim.Genes

1:Name_off_pattient 2:LOH	3:4	4:126303972	5:128296809	6:1992.838	7:437	8:4570	9:q28.1	10:q28.2	11:8	12:6	13:INTU, SLC25A31, HSPA4L, PLK4, MFSD8, ABHD18, LARP1B, PGRMC2	14:INTU (610621), SLC25A31 (610796), HSPA4L (619077), PLK4 (605031), MFSD8 (611124), PGRMC2 (607735)

5-The program will finish its work by generating a document!
You need to write a convenient place for you to generate this file.

My file is saved to the desktop and is called imprinting_analysis 


Thank you all for your attention. 

With you was 
# Vasin Kirill Sergeevich, MD., Phd. 
My contacts:

drvasinks@gmail.com

telegram @DrKvasin

https://www.linkedin.com/in/kirill-vasin-ba654a104/
>>>>>>> 20006a7cabf8ed3727c5a02b6db90339c64d929d
