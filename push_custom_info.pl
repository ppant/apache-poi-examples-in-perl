#!/usr/local/bin/perl    
	# Call to push custom info in word documents	
	# Handle DOC and SLS extesnions 
	
    if (defined($content_type) && ($content_type eq "application/x-msword" || $content_type eq "application/vnd.ms-excel" || $content_type eq "application/vnd.ms-powerpoint") ) {
	my $path = $POI_temp_path."poi_view_tmp.doc";
	my $buffer_temp;
	open (MSHANDLE , ">".$path) or die "Can't open: $!\n";
	print MSHANDLE $buffer;
	close(MSHANDLE);
	my $alu = POI->new();
	my $docname = $alu->PushCustomProperties($path,"Test_doc_file","ppant","2.0", now());	
	open (MSHANDLE , $path) or die "Can't open: $!\n";
	binmode(MSHANDLE);
	read(MSHANDLE, $buffer_temp, -s $path);
	close(MSHANDLE);
	$buffer = $buffer_temp;
	unlink($path);
	}
	
	
	
	
	