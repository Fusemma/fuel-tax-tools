# Forms templates

What it does: Word templates for the certificates redone by hand at work.
The 637 notification certificate (Notification Certificate - TEMPLATE.docx,
shipped) and the federal excise exemption certificate (Exemption
Certificate - TEMPLATE.docx, Models M and P in one form, built and
tested). Every fill-in spot is a click-and-type field with a gray hint;
signature and date stay blank for ink. word_fields.py writes the fields;
build_637_template.py and build_exemption_template.py make the files;
render_proof.py makes a picture for a look-over. The wording is in
build_637_template.py and follows IRS Pub 510 (p510.pdf, beside this
file).

At work: download the .docx, fill the fields, print. The Word cheat sheet
beside this file says how the fields behave.

Open: a look-over of the exemption certificate and the Word cheat sheet,
then it ships.
