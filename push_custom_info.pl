#!/usr/local/bin/perl    
	# Call to push custom info in word documents	
	use POI;	# This module will deal with doc type
	use POIPlus; # This module will deal with docx type	
	# Make sure to give full path of the doc and docx file
	my $doc_path = "doc_test_file.doc";
	my $docx_path = "docx_test_file.docx";
	# Make a object
	my $poi_obj = POI->new();
	my $poiplus_obj = POIPlus->new();
	# Call push properties routine which is inside POI.pm and written in Java a Java module
	my $doc_name = $poi_obj->PushProperties($doc_path,"Test_doc_file","Test file doc type extension","2.0","ppant");
	print "Properties written into doc file ".$doc_path."\n";	
	# Call push properties routine which is inside POIPlus.pm and written in a Java module
	my $docx_name =  $poiplus_obj->PushProperties_docx($docx_path,"Test_docx_file","Test file docx type extension","1.0","ppant");
	print "Properties written into docx file ".$docx_path."\n";	
	
	
	
	
	
	
	
