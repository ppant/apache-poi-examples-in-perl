#!/usr/local/bin/perl    
	use POI;
	# Call to push custom info in word documents	
	#!/usr/local/bin/perl    
	use POI;	
	# Make sure to give full path of the doc file
	my $path = "doc_test_file.doc";
	# Make a POI object
	my $poi_obj = POI->new();
	# Call push properties routine which is inside POI.pm and written in Java
	my $docname = $poi_obj->PushProperties($path,"Test_doc_file","Test file doc type extension","2.0","ppant");
	
	
	
	
	
	
