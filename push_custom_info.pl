#!/usr/local/bin/perl    
	use POI;
	# Call to push custom info in word documents	
	# Handle DOC, XLS  and PPT extension 
	# Todo
	#my $content_type;
    #if (defined($content_type) && ($content_type eq "application/x-msword" || $content_type eq "application/vnd.ms-excel" || $content_type eq "application/vnd.ms-powerpoint") ) {
	# Make sure to give full path of the doc file
	my $path = "doc_test_file.doc";
	# Make a POI object
	my $alu = POI->new();
	# Time to call push properties routine which is inside POI.pm and written in Java
	my $docname = $alu->PushProperties($path,"Test_doc_file","Test file doc type extension","2.0","ppant");	
	
	
	
	
	
	
