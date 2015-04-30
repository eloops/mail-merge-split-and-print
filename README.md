# mail-merge-split-and-print
Prints specified sections of a mail merged document to a target printer

When you mail merge a document, Word outputs a single document, with each set separated by a Section.

This macro will send each section to the specified local print queue (network ones work too I think). The client in this case wanted each set stapled separately.

To set this up:
* Enable the Developer toolbar in Word
* Open the Macro editor
* Create a new macro and edit it
* Paste in this code & adjust to suit your own environment
* Set up local print queue with whatever defaults you need per set (stapled, hole punched, whatever)
* Run mail merge to output a document (NOT merge & print)
* Run macro on merged document, it should send a pile of print jobs VERY quickly.

If your base document has sections already in it, you may need to adjust the 'counter' variables. I'd recommend testing on a small subset of recipients before doing it on a larger scale set.
