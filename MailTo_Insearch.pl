#! perl shebang line
# This script is used to automate the return of marking sheets and feedback to students
# assumed directory structure is
# Tutorial Group
#	Student Number
# This script will work with any number of tutorial groups - edit parameters passed to batchMail to increase or decrease number of tutorials
# Update the text files T1.txt, T2.txt etc with the student numbers for each tutorial each semester

#Global variables for script - change as required

$email_title = 'INFO322 assignment 1 Part A results'; #subject line of the email
$msg_content = 'Here is your feedback sheet for INF0322 assignment 1 part A.'; #body of the email
$filename = 'INFO322Marking.doc'; #file to attach - must be located in [tutorialGroup]\[studentNumber] directory

#Constants - all student email addresses alias to [studentNumber]@uts.edu.au
$insearch = '@uts.edu.au'; #email domain for insearch students

# Remove or add any tutorials as required
# Make sure a matching txt file exists in the same diectory as this script
&batchMail(T4); #Here I am sending feedback to 2 tutorial groups 

#DO NOT CHANGE ANY CODE BELOW HERE!!!!

# This subroutine is used to send mail to each tutorial group, following the directory structure
# Tutorial Group
#	Student Number
sub batchMail
{
	my(@tutorials) = @_;
	#create the directories for each tute and each student in each tute
	foreach my $tutorial (@tutorials)
	{
		chomp($tutorial);
		&sendMail($tutorial, "$tutorial.txt");
	} 
}

 # This subroutine is used to create an email to each student using Outlook, attaching the marking sheet 
 # Tutorial Group
 #	Student Number
 sub sendMail 
 {
 	my($tute, $dat_file) = @_;
 	use Mail::Outlook;
	# start with a folder
	my $outlook = new Mail::Outlook('Inbox');

 	open(DAT, $dat_file) || die("Could not open file!");
 	my @students = <DAT>;
 	foreach my $student (@students)
 	{
  		chomp($student);
  		# create a message for sending
		my $message = $outlook->create();
		$message->To("$student$insearch");
		$message->Subject("$email_title");
		$message->Body("$msg_content");
		use Cwd;
		my $dir = getcwd; # need path to the root directory
		my $slash = '/';
		my $backslash = "\\";	
		$dir =~ s/$slash/$backslash/g; #replace unix \ with / for DOS - really ugly but works :)
		$message->Attach("$dir\\$tute\\$student\\$filename");
		$message->send || die("Could not send email!");
		  
        	print "$student email sent\n";
        	sleep(10); #wait for process to complete
 	} 
 	close(DAT);
 	print "$tute emails sent\n";
}
