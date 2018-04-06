#!/usr/bin/perl

$| = 1;        # AutoFlush to ensure that stdout & stderr are in sync
use 5.006_001;
use strict;
use warnings;
use Cwd;       # Get pathname of current working directory
use open ':std', ':encoding(UTF-8)';

use Win32::OLE qw( in );
Win32::OLE -> Option( CP => Win32::OLE::CP_UTF8 );

# =================================
# Global Variables
my %AllPlaylist;
my @RootList = ();
my $RootCnt  = 0;
my $RootPath = "";
#my @AllSong  = "";
my $_DEBUG_ = 0;

# Define the subroutine to call when the user press CTRL-C
$SIG {"INT"} = \&CleanExit; # "INT" indicates "Interrupt" signal.

# =================================
# Debug Message
sub DebugMsg
{
   print STDOUT "-- Debug: @_\n" unless( $_DEBUG_ eq 0 );
   #if( $_DEBUG_ eq 1 )
   #{
   #   print (Win32::OLE -> LastError()) ||  "-- Debug: @_\n";
   #}
}

# =================================
# CTRL-C Interrupt handler
sub CleanExit()
{
   print STDOUT "\n******************************************\n";
   print STDOUT "  Exiting Playlist Creation...\n";
   print STDOUT "******************************************\n";
   exit 0;
}

# =================================
# Subroutine to parse through the files and directories within the directory
sub ParseCurrentDirectory
{
   my( $ThisDir ) = @_;
   my $CurrentMp3List = "";

   # Setup the Win32 object
   DebugMsg( "--- Testing \"$ThisDir\" ---" );
   my $Win32Obj    = Win32::OLE -> new( 'Scripting.FileSystemObject' );
   my $ThisFolder  = $Win32Obj -> GetFolder( $ThisDir );
   my $AllFileList = $ThisFolder -> {Files};
   my $AllDirList  = $ThisFolder -> {SubFolders};

   # Go through the sub-directories of the current folder
   foreach my $EachDir( in $AllDirList )
   {
      # Get the directory name and filter out the directories starting with "."
      my $ThisDirName = $EachDir -> {Name};
      next if( $ThisDirName =~ /^\./ );
      
      # Parse this value into this function recursively
      my $FullPath = "${ThisDir}/${ThisDirName}";
      if( $ThisDir eq $RootPath )
      {
         DebugMsg( "I'm in Root Path" );
         @RootList = ();
      }
      if( !($ThisDir eq $RootPath) )
      {
         $RootCnt++;
         $RootList[$RootCnt] = "";
      }
      ParseCurrentDirectory( $FullPath );
      if( !($ThisDir eq $RootPath) )
      {
         $RootList[$RootCnt] = "";
         $RootCnt--;
      }
   }

   # Go through the list of files (not directory)
   foreach my $EachFile( in $AllFileList )
   {
      # Get the filename and filter out the non MP3 extension files.
      my $ThisFilename = $EachFile -> {Name};
      next if( $ThisFilename !~ /\.mp3$/ );

      # Save the full path of the file without the root
      my $FullPathName = "${ThisDir}/${ThisFilename}";
      $FullPathName =~ s/\Q${RootPath}\E\///g;
      DebugMsg( "Saving $FullPathName into List" );
      $CurrentMp3List .= "$FullPathName\n";
   }
   
   # Do something only if the current list is not empty
   #if( $CurrentMp3List !~ "" )
   #{
      my $PlaylistFilename = "${ThisDir}";
      $PlaylistFilename =~ s/\Q${RootPath}\E//g;
      $PlaylistFilename =~ s/\s+//g;
      $PlaylistFilename =~ s/\///g;
      
      #DebugMsg( "List is not empty." );
      # Save the list in all directories containing the current one
      foreach my $idx( 0 .. ($RootCnt - 1) )
      {
         $RootList[$idx] .= $CurrentMp3List;
      }
      DebugMsg( "Saving \"$PlaylistFilename\" with \"$RootList[$RootCnt]\n$CurrentMp3List\"" );
      $AllPlaylist{$PlaylistFilename}  = $RootList[$RootCnt];
      $AllPlaylist{$PlaylistFilename} .= $CurrentMp3List;
      $AllPlaylist{"AllSongs"} .= $CurrentMp3List;
   #}
}

## =================================
## Subroutine to parse through the files and directories within the directory
#sub FileDirParser
#{
#   my( $ThisDir ) = @_;
#   my @FileList = "";
#
#   undef @FileList;
#   DebugMsg( "---\nTesting \"$ThisDir\"" );
#   print STDOUT "Directory: $ThisDir\n" unless( $_DEBUG_ eq 1 );
#
#   # Setup the Win32 object
#   my $Win32Obj    = Win32::OLE -> new( 'Scripting.FileSystemObject' );
#   my $ThisFolder  = $Win32Obj -> GetFolder( $ThisDir );
#   my $AllFileList = $ThisFolder -> {Files};
#   my $AllDirList  = $ThisFolder -> {SubFolders};
#
#   # Go through the list of files (not directory)
#   foreach my $EachFile( in $AllFileList )
#   {
#      # Get the filename and filter out the non MP3 extension files.
#      my $ThisFilename = $EachFile -> {Name};
#      next if( $ThisFilename !~ /\.mp3$/ );
#
#      # Save the full path of the file without the root
#      my $FullPathName = "${ThisDir}/${ThisFilename}";
#      $FullPathName =~ s/\Q${RootPath}\E\///g;
#      DebugMsg( "Saving $FullPathName into List" );
#      push @FileList, "$FullPathName\n";
#      push @AllSong, "$FullPathName\n";
#   }
#
#   # Save the name into the playlist
#   if( @FileList )
#   {
#      my $PlaylistFilename = "${ThisDir}.m3u";
#      $PlaylistFilename =~ s/\Q${RootPath}\E//g;
#      $PlaylistFilename =~ s/\s+//g;
#      $PlaylistFilename =~ s/\///g;
#      
#      # Save the list into the playlist
#      my $PlaylistLocation = "${RootPath}/${PlaylistFilename}";
#      open( my $INPUTF, ">$PlaylistLocation" );
#      print $INPUTF @FileList;
#      close( $INPUTF );
#      DebugMsg( "Created \"$PlaylistFilename\" with @FileList" );
#   }
#
#   # Go through the sub-directories of the current folder
#   foreach my $EachDir( in $AllDirList )
#   {
#      # Get the directory ame and filter out the directories starting with "."
#      my $ThisDirName = $EachDir -> {Name};
#      next if( $ThisDirName =~ /^\./ );
#      
#      # Parse this value into this function recursively
#      my $FullPath = "${ThisDir}/${ThisDirName}";
#      FileDirParser( $FullPath );
#   }
#}

# =================================
# Main Subroutine
sub main()
{
   # Get the full path of the current directory
   $RootPath = cwd();
   print STDOUT "This is the root directory: \"$RootPath\"\n";

   # Clear all playlist in the current directory
   foreach my $delfiles( glob "*.m3u" )
   {
      unlink $delfiles;
   }

#   # Open the current directory for listings
#   FileDirParser( "$RootPath" );
#   
#   # Save all songs into a single playlist
#   my $AllSongName = "AllSongs.m3u";
#   open( my $INPUTF, ">$AllSongName" );
#   print $INPUTF @AllSong;
#   close( $INPUTF );

   # Open the current directory for listings
   ParseCurrentDirectory( "$RootPath" );

   # Create all the playlists
   foreach my $listnames( keys %AllPlaylist )
   {
      next if( $listnames eq "" );
      next if( $AllPlaylist{$listnames} eq "" );

      my $PlaylistName = "$listnames.m3u";
      DebugMsg( "Generating \"$PlaylistName\"..." );
      open( my $INPUTF, ">$PlaylistName" );
      print $INPUTF $AllPlaylist{$listnames};
      close( $INPUTF );
   }

   # Quit the program
   CleanExit();
}

# =================================
# Execute Main Subroutine
main();
