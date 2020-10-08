#!/bin/bash

# history_of_file
#
# Outputs the full history of a given file as a sequence of
# logentry/diff pairs.  The first revision of the file is emitted as
# full text since there's not previous version to compare it to.

function history_of_file() {
    url=$1 # current url of file
    svn log -q $url | grep -E -e "^r[[:digit:]]+" -o | cut -c2- | sort -n | {

#       first revision as full text
        echo
        read r
        svn log -r$r $url@HEAD > $url-FullSvnHistory.txt
        svn cat -r$r $url@HEAD >> $url-FullSvnHistory.txt
        echo

#       remaining revisions as differences to previous revision
        while read r
        do
            echo
            svn log -r$r $url@HEAD >> $url-FullSvnHistory.txt
            svn diff -c$r $url@HEAD >> $url-FullSvnHistory.txt
            echo
        done
    }
}

history_of_file $1