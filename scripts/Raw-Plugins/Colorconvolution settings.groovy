//Low contrast 1 - For exocrine images that are not lookin nice n shit yknow?

setColorDeconvolutionStains(
    '{"Name" : "H-DAB-low contrast-2", ' +
    '"Stain 1" : "Hematoxylin", ' +
    '"Values 1" : "0.7778767882082012 0.5475636430647997 0.3083533024964999", ' +
    '"Stain 2" : "miR-155", ' +
    '"Values 2" : "0.3455742195039925 0.84422186476732 0.40971685572233324", ' +
    '"Background" : "255 255 255"}'
);


// Low contrast 2 - For when you see the nuclei more clearly and don't need to be too aggressive on the threshold 

setColorDeconvolutionStains(
    '{"Name" : "H-DAB estimated", ' +
    '"Stain 1" : "Hematoxylin", ' +
    '"Values 1" : "0.7989797507243244 0.5482931059008459 0.24699398363948147", ' +
    '"Stain 2" : "miR-155", ' +
    '"Values 2" : "0.351847155609121 0.928596612388845 0.11794876239169397", ' +
    '"Background" : "255 255 255"}'
);