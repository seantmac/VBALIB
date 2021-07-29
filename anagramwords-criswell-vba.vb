Option Compare Database
Option Explicit

'/*
'Wordplay Version 7.22         03-20-96
'Written by Evans A Criswell at the University of Alabama in Huntsville
'03-20-96 Fixed a small memory allocation problem.  In a couple of places,
'   the amount allocated to hold character strings was not taking the
'   space to store the null into account.  This bug has only affected
'   a couple of people.
'09-11-95 In the anagramr7 function, I check the product of the maximum
'   "levels deep" remaining and the length of the longest candidate
'   word.  If this product is less than the length of the string
'   passed in, a "dead end" condition exists.  This makes the program
'   run significantly faster for longer strings if the maximum
'   depth option is used.
'08-21-94 Added "wordfile from stdin" option using "-f -"
'   Fixed "4" bug.  Digits in a string disqualify the string.
'   Vowel-check override option added.
'   Starting word ("w" option) checked to see if it's an anagram
'   of the initial string.
'08-16-94 Used integer masks representing which letters appear in each
'   word, allowing extraction checking to be checked quickly for
'   failure in the anagramr7 routine.  Result:  the program has
'   been 4 to 5 times faster.
'08-14-94 Made the program much more memory efficient.  Instead of calling
'   malloc for each word in the candidate word list and in the key
'   list, a contiguous block of memory was allocated to hold the
'   words.  The block is realloc'ed if it needs to be increased as
'   the words are read in.  After the words are packed into the
'   block, the pointers are allocated and are pointed to the
'   appropriate places (beginnings of words) in the block, so the
'   rest of the program works with no modification.  Two gigantic
'   arrays that weren't being used were eliminated.  The word length
'   index arrays are now made to be the size of the longest word
'   instead of MAX_WORDS.  In fact, MAX_WORDS is now obsolete.
'07-14-94 Added "silent" option.
'06-03-94 Added "#include <ctypes.h>" so it would work on BSD/386 .  Thanks
'         to mcintyre@io.com (James Michael Stewart) for reporting the bug.
'05-26-94 Fixed command-line parsing bug.
'05-25-94 Eliminated redundant permutations.  Added option to specify a
'   word to appear in anagrams.  Added maximum depth option (number
'   of words, maximum, to appear in an anagram).
'05-24-94 Added option so user could specify whether to allow anagrams
'   with adjacent duplicate words like "A A" or "DOG DOG".
'05-16-94 Made a second copy of the word list and sorted each word's
'   letters alphabetically and sorted this list of keys alphabetically.
'   Modified the recursive algorithm to use the new index. (Ver 6.00)
'05-16-94 Another little bug fix.  Someone found that, on their machine,
'   if there are no candidate words loaded for the string being
'         anagrammed, it causes an error when malloc gets passed a zero
'         value for the amount to allocate.
'05-13-94 Tiny bug fix.  Just a small bug that never actually caused a
'   crash, but very well could have if it had wanted to.  :-)
'04-25-94 Speed increase.  If exts indicates extraction was impossible,
'   continue (try next word) instead of executing rest of loop body.
'04-21-91 Ron Gregory found a simple bug that has been in all the C
'   versions (4.00 through 5.20).  In the one-word anagram
'   section, a less than should have been a less than or equal to.
'   A simple fencepost error.  The recursive anagram procedure had
'   a similar problem.  A severe error was fixed in the version
'   5.20 read routine which caused the program not to read the
'   wordfile correctly if the entries were lowercase.
'04-17-94 Since this program, since it was ported to C, is command-line
'   based, and only anagrams one string, it is not necessary to
'   store the wordlist internally.  Unnecessary words are weeded
'   out as the list is being read, using the "extract" routine.
'   I can't believe I didn't think of using that routine for that
'   purpose sooner.  That means pass1 and pass2 are obsolete.
'04-14-94 Changed the "extract" function to use pointers instead of
'   array notation.  Under some compilers, this may nearly double
'   the execution speed of the recursive anagram procedure.  On
'   other compilers, it may make no difference at all.
'04-11-94 Added the minimum and maximum candidate word length options
'   that were available in version 3.00 when the program was
'   interactive.  This helps to narrow down the word list and
'   eliminate a lot of short words when anagramming long strings.
'11-30-93 Fixed a bug that Versions 5.00 and 5.01 had.  If there were
'   no words in the candidate word list with the same length as
'   the string passed to anagramr, the string passed to anagramr
'   would not be anagrammed, causing many possible anagrams to
'   be missed.
'11-08-93 Eliminated anagrams consisting of the same word occurring
'         multiple times in a row, such "IS IS ...", since interesting
'         anagrams rarely contain such repetitions. (Version 5.01)
'11-08-93 Debug print statements commented and output cleaned up.
'         Version 5.00 completed.  It is currently not known which is
'   always faster:  the old iterative 2 and 3 word anagram options
'   or the recursive algorithm.  All the options from version 4.00
'   are still in the program.
'11-07-93 Recursive algorithm working!
'11-03-93 Added code to index the candidate word list by number of vowels
'   per word. (Beginning of 5.00 Alpha)  Never used in Version 5.00,
'   but the code is there for future use.
'05-25-93 Three word anagramming capability ported and added.
'04-30-93 The big port from FORTRAN 77 to ANSI C.  No longer interactive.
'   Instead, arguments are taken from the command line.
'   (Everything working except three-word anagrams and all command
'   line options not yet implemented)
'Version 4.00 is the first version to be implemented in C.  All previous
'versions were written in FORTRAN 77.
'Note:  There was no version 5.12.  It was called 5.20 instead.
'Version 7.22  03-20-96  Bug fix.
'Version 7.21  09-11-95  Speed increase.
'Version 7.20  08-21-94  Wordfile from stdin capability, bug fixes.
'Version 7.11  08-16-94  Speed increase.
'Version 7.10  08-14-94  Program uses much less memory.
'Version 7.02  07-14-94  Silent option.
'Version 7.01  06-03-94  Portability problem fixed.  ctypes.h needed .
'Version 7.00  05-26-94  Redundant permutations eliminated. Several refinements.
'Version 6.00  05-17-94  Huge speed increase.
'Version 5.24  05-16-94  Bug fix.
'Version 5.23  05-13-94  Tiny bug fix.
'Version 5.22  04-25-94  Speed increase.
'Version 5.21  04-21-94  Bug fixes.
'Version 5.20  04-17-94  Faster program initialization.  Far less memory used.
'Version 5.11  04-14-94  Slight speed increase with some compilers
'Version 5.10  04-11-94  Minimum, maximum candidate word length again
'        available.  (First time available in the C versions).
'Version 5.02  11-30-93  Bug fix.
'Version 5.01  11-08-93  Optimization to eliminate multiple occurrences
'                        of a particular word in a row.
'Version 5.00  11-08-93  Recursive algorithm added
'Version 4.00  04-30-93  Ported to C.  Became non-interactive and more
'        suitable for UNIX environment
'Version 3.00  12-16-91  Indexing improvements.  Huge speed increase
'Version 2.10  04-16-91  Options and help added
'Version 2.00  04-12-91  Three word anagrams added
'Version 1.11  04-11-91  Bug fixes and cleanups
'Version 1.10  04-03-91  Pass 2 word filter added.  Huge speed increase.
'Version 1.00  03-29-91  One and two word anagrams
'*/
'
'#include <stdlib.h>
'#include <stdio.h>
'#include <string.h>
'#include <ctype.h>
'#define max(A, B) ((A) > (B) ? (A) : (B))
'#define min(A, B) ((A) < (B) ? (A) : (B))
'#define DEFAULT_WORD_FILE "WORDS721.TXT"
'#define WORDBLOCKSIZE 4096
'#define MAX_WORD_LENGTH 128
'#define SAFETY_ZONE MAX_WORD_LENGTH + 1
'#define MAX_ANAGRAM_WORDS 32
'#define MAX_PATH_LENGTH 256


Public Const gSUCCESS      As Integer = 0
Public Const gFAIL         As Integer = -1

Public Const DEFAULT_WORD_FILE As String = "WORDS721.TXT"
Public Const WORDBLOCKSIZE     As Integer = 4096
Public Const MAX_WORD_LENGTH   As Integer = 128
Public Const SAFETY_ZONE       As Integer = 129   '(MAX_WORD_LENGTH + 1)
Public Const MAX_ANAGRAM_WORDS As Integer = 32
Public Const MAX_PATH_LENGTH   As Integer = 256


'*&*&
Dim findx1() As Integer
[26
UnknownDim findx2() As Integer
[26
Unknown
    
'*&*&
    Private Overloads Function uppercase(ByVal s As Char) As Char
    End Function
    
    Private Overloads Function alphabetic(ByVal s As Char) As Char
    End Function
    
    Private Overloads Function numvowels(ByVal s As Char) As Integer
    End Function
    
    Private Overloads Sub anagramr7(ByVal s As Char, ByVal Star As Char, ByVal , As accum, ByVal minkey As Integer, ByVal level As Integer)
    End Sub
    
    Private Overloads Function extract(ByVal s1 As Char, ByVal s2 As Char) As Char
    End Function
    
    Private Overloads Function intmask(ByVal s As Char) As Integer
    End Function
'*&*&


' char   *uppercase (char *s);
' char   *alphabetic (char *s);
' int     numvowels (char *s);
' void    anagramr7 (char *s, char **accum, int *minkey, int *level);
' char   *extract (char *s1, char *s2);
' int     intmask (char *s);
' char  **words2;  /* Candidate word index (pointers to the words) */
' char   *words2mem;  /* Memory block for candidate words  */
' char  **words2ptrs; /* For copying the word indexes */
' char  **wordss;    /* Keys */
' char   *keymem;     /* Memory block for keys */
' int    *wordsn;    /* Lengths of each word in words2 */
' int    *wordmasks; /* Mask of which letters are contained in each word */
' int     ncount;    /* Number of candidate words */
' int     longestlength; /*  Length of longest word in words2 array */
' char    largestlet;
' int     rec_anag_count;  /*  For recursive algorithm, keeps track of number
' 			 of anagrams fond */
' int     adjacentdups;
' int     specfirstword;
' int     maxdepthspec;
' int     silent;
' int     max_depth;
' int     vowelcheck;
' int    *lindx1; 
' int    *lindx2;
' int     findx1[26];
' int     findx2[26];


Private words2mem      As  Characters
Private keymem         As  Characters
Private wordsn         As  Integer
Private wordmasks      As  Integer
Private ncount         As  Integer
Private longestlength  As  Integer
Private largestlet     As  Characters
Private rec_anag_count As Integer
Private adjacentdups   As Integer
Private specfirstword  As Integer
Private maxdepthspec   As Integer
Private silent         As Integer
Private max_depth      As Integer
Private vowelcheck     As Integer
Private lindx1         As Integer
Private lindx2         As Integer
    

'' C-VERSION:   #define max(A, B) ((A) > (B) ? (A) : (B))
'' C-VERSION:   #define min(A, B) ((A) < (B) ? (A) : (B))

Private Function max(x as variant, y As Variant) As Variant
   max = IIf(x > y, x, y)
End Function

private Function min(x as Variant, y As Variant) As Variant
    min = IIf(x < y, x, y)
End Function


    Private Function main(ByVal argc As Integer, ByVal argv() As Char) As Integer
        Dim word_file_ptr As file
        Dim buffer() As char
        MAX_WORD_LENGTH
        Dim ubuffer() As char
        MAX_WORD_LENGTH
        Dim alphbuffer() As char
        MAX_WORD_LENGTH
        Dim initword() As char
        MAX_WORD_LENGTH
        Dim remaininitword() As char
        MAX_WORD_LENGTH
        Dim word_file_name() As char
        MAX_PATH_LENGTH
        Dim first_word() As char
        MAX_WORD_LENGTH
        Dim u_first_word() As char
        MAX_WORD_LENGTH
        Dim tempword() As char
        MAX_WORD_LENGTH
        Dim ilength As Integer
        Dim size As Integer
        Dim gap As Integer
        Dim switches As Integer
        Dim iholdn As Integer
        Dim chold As char
        Dim wholdptr As char
        Dim curlen As Integer
        Dim curpos As Integer
        Dim curlet As char
        Dim icurlet As Integer
        Dim recursiveanag As Integer
        Dim listcandwords As Integer
        Dim wordfilespec As Integer
        Dim firstwordspec As Integer
        Dim maxcwordlength As Integer
        Dim mincwordlength As Integer
        Dim iarg As Integer
        Dim keyi As Integer
        Dim keyj As Integer
        Dim Star As char
        accum
        Dim level As Integer
        Dim minkey As Integer
        Dim leftover() As char
        MAX_WORD_LENGTH
        Dim w2size As Integer
        Dim w2memptr As char
        Dim w2offset As Integer
        Dim keymemptr As char
        Dim keyoffset As Integer
        Dim no() As char
3
        "no"
        Dim yes() As char
4
        "yes"
        Dim fileinput As Integer
        Dim hasnumber As Integer
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        If (argc < 2) Then
            fprintf(stderr, "Wordplay Version 7.22  03-20-96, 1991   by Evans A Criswell"& vbLf)
            fprintf(stderr, "University of Alabama in Huntsville     criswell@cs.uah.edu"& vbLf& vbLf)
            fprintf(stderr, "Usage:  ")
            fprintf(stderr, "wordplay string_to_anagram [-slxavnXmXdX] [-w word] ", "[-f word_file]"& vbLf& vbLf)
            fprintf(stderr, "Capital X represents an integer."& vbLf& vbLf)
            fprintf(stderr, "s  = silent operation (no header or line numbers)"& vbLf)
            fprintf(stderr, "l  = print candidate word list"& vbLf)
            fprintf(stderr, "x  = do not generate anagrams (useful with l option)"& vbLf)
            fprintf(stderr, "a  = multiple occurrences of a word in an anagram OK"& vbLf)
            fprintf(stderr, "v  = allow words with no vowels to be considered"& vbLf)
            fprintf(stderr, "nX = candidate words must have n characters minimum"& vbLf)
            fprintf(stderr, "mX = candidate words must have m characters maximum"& vbLf)
            fprintf(stderr, "dX = limit anagrams to d words"& vbLf& vbLf)
            fprintf(stderr, "w word = word to start anagrams"& vbLf)
            fprintf(stderr, "f file = word file to use (""-f -"" for stdin)"& vbLf& vbLf)
            fprintf(stderr, "Suggestion:  Run ""wordplay trymenow"" ", " to get started."& vbLf)
            exit(-1)
        End If
        
        strcpy(word_file_name, DEFAULT_WORD_FILE)
        recursiveanag = 1
        listcandwords = 0
        wordfilespec = 0
        firstwordspec = 0
        specfirstword = 0
        silent = 0
        vowelcheck = 1
        maxdepthspec = 0
        maxcwordlength = MAX_WORD_LENGTH
        mincwordlength = 0
        max_depth = MAX_ANAGRAM_WORDS
        iarg = 1
        
        While (iarg < argc)
            If (wordfilespec = 1) Then
                strcpy(word_file_name, argv(iarg))
                iarg = (iarg + 1)
                wordfilespec = 0
                'TODO: Warning!!! continue If
            End If
            
            If (firstwordspec = 1) Then
                strcpy(first_word, argv(iarg))
                iarg = (iarg + 1)
                firstwordspec = 0
                'TODO: Warning!!! continue If
            End If
            
            If (argv(iarg)(0) = Microsoft.VisualBasic.ChrW(45)) Then
                If (CType(strlen(argv(iarg)),Integer) > 1) Then
                    i = 1
                    
                    While (i < CType(strlen(argv(iarg)),Integer))
                        Select Case (argv(iarg)(i))
                            Case Microsoft.VisualBasic.ChrW(97)
                                adjacentdups = 1
                            Case Microsoft.VisualBasic.ChrW(108)
                                listcandwords = 1
                            Case Microsoft.VisualBasic.ChrW(102)
                                wordfilespec = 1
                            Case Microsoft.VisualBasic.ChrW(120)
                                recursiveanag = 0
                            Case Microsoft.VisualBasic.ChrW(115)
                                silent = 1
                            Case Microsoft.VisualBasic.ChrW(118)
                                vowelcheck = 0
                            Case Microsoft.VisualBasic.ChrW(119)
                                firstwordspec = 1
                                specfirstword = 1
                            Case Microsoft.VisualBasic.ChrW(109)
                                maxcwordlength = 0
                                i = (i + 1)
                                
                                While ((argv(iarg)(i) >= Microsoft.VisualBasic.ChrW(48)) _
                                            AndAlso (argv(iarg)(i) <= Microsoft.VisualBasic.ChrW(57)))
                                    maxcwordlength = ((maxcwordlength * 10) _
                                                + (CType(argv(iarg)(i++),Integer) - CType(Microsoft.VisualBasic.ChrW(48),Integer)))
                                    
                                End While
                                
                                i = (i - 1)
                            Case Microsoft.VisualBasic.ChrW(110)
                                i = (i + 1)
                                
                                While ((argv(iarg)(i) >= Microsoft.VisualBasic.ChrW(48)) _
                                            AndAlso (argv(iarg)(i) <= Microsoft.VisualBasic.ChrW(57)))
                                    mincwordlength = ((mincwordlength * 10) _
                                                + (CType(argv(iarg)(i++),Integer) - CType(Microsoft.VisualBasic.ChrW(48),Integer)))
                                    
                                End While
                                
                                i = (i - 1)
                            Case Microsoft.VisualBasic.ChrW(100)
                                maxdepthspec = 1
                                max_depth = 0
                                i = (i + 1)
                                
                                While ((argv(iarg)(i) >= Microsoft.VisualBasic.ChrW(48)) _
                                            AndAlso (argv(iarg)(i) <= Microsoft.VisualBasic.ChrW(57)))
                                    max_depth = ((max_depth * 10) _
                                                + (CType(argv(iarg)(i++),Integer) - CType(Microsoft.VisualBasic.ChrW(48),Integer)))
                                    
                                End While
                                
                                i = (i - 1)
                            Case Else
                                fprintf(stderr, "Invalid option: ""%c"" - Ignored"& vbLf, argv(iarg)(i))
                        End Select
                        
                        i = (i + 1)
                        
                    End While
                    
                End If
                
                iarg = (iarg + 1)
            Else
                strcpy(initword, uppercase(argv(iarg)))
                iarg = (iarg + 1)
            End If
            
            
        End While
        
        If (silent = 0) Then
            printf ("Wordplay Version 7.22  03-20-96, 1991   by Evans A Criswell" & vbLf)
            printf("University of Alabama in Huntsville     criswell@cs.uah.edu"& vbLf& vbLf)
        End If
        
        If (silent = 0) Then
            printf ("" & vbLf)
            printf("Candidate word list :  %s"& vbLf, (listcandwords = 0))
            'TODO: Warning!!!, inline IF is not supported ?
            'TODO: Warning!!!! NULL EXPRESSION DETECTED...
            
            printf("Anagram Generation  :  %s"& vbLf, (recursiveanag = 0))
            'TODO: Warning!!!, inline IF is not supported ?
            'TODO: Warning!!!! NULL EXPRESSION DETECTED...
            
            printf("Adjacent duplicates :  %s"& vbLf, (adjacentdups = 0))
            'TODO: Warning!!!, inline IF is not supported ?
            'TODO: Warning!!!! NULL EXPRESSION DETECTED...
            
            printf("Vowel-free words OK :  %s"& vbLf& vbLf, (vowelcheck = 0))
            'TODO: Warning!!!, inline IF is not supported ?
            'TODO: Warning!!!! NULL EXPRESSION DETECTED...
            
            printf("Max anagram depth   :  %d"& vbLf, max_depth)
            printf("Maximum word length :  %d"& vbLf, maxcwordlength)
            printf("Minimum word length :  %d"& vbLf& vbLf, mincwordlength)
            If specfirstword Then
                printf("First word          :  ""%s"""& vbLf, first_word)
            End If
            
            printf("Word list file      :  ""%s"""& vbLf, word_file_name)
            printf("String to anagram   :  ""%s"""& vbLf, initword)
            printf ("" & vbLf)
        End If
        
        strcpy(tempword, alphabetic(initword))
        strcpy(initword, tempword)
        ilength = CType(strlen(initword),Integer)
        size = ilength
        gap = size
        
        Do Until ((switches <> 0) _
                    Or (gap <> 1))
            gap = max(((gap * 10) _
                            / 13), 1)
            switches = 0
            i = 0
            Do While (i _
                        < (size - gap))
                j = (i + gap)
                If (initword(i) > initword(j)) Then
                    chold = initword(i)
                    initword(i) = initword(j)
                    initword(j) = chold
                    switches = (switches + 1)
                End If
                
                i = (i + 1)
            Loop
            
            
        Loop
        
        If specfirstword Then
            strcpy(u_first_word, uppercase(first_word))
            strcpy(remaininitword, extract(initword, u_first_word))
            If (remaininitword(0) = Microsoft.VisualBasic.ChrW(48)) Then
                fprintf(stderr, "Specified first word ""%s"" cannot be extracted ", "from initial string ""%s"""& vbLf, u_first_word, initword)
                exit(1)
            End If
            
            If (StrLen(remaininitword) = 0) Then
                If (silent = 0) Then
                    printf ("Anagrams found:" & vbLf)
                    printf("     0.  %s"& vbLf, u_first_word)
                Else
                    printf("%s"& vbLf, u_first_word)
                End If
                
                exit(0)
            End If
            
        End If
        
        w2size = WORDBLOCKSIZE
        If (ctype(malloc(w2size, Star, sizeof, char), char) = ctype(Null, char)) Then
            fprintf(stderr, "Insufficient memory; malloc returned NULL."& vbLf)
            exit(-1)
        End If
        
        If (silent = 0) Then
            printf(""& vbLf&"Initializing.  Please wait while words are being loaded"& vbLf)
            printf ("and unnecessary words are being filtered out ..." & vbLf)
        End If
        
        If (strcmp(word_file_name, "-") = 0) Then
            fileinput = 0
            word_file_ptr = StdIn
        ElseIf (fopen(word_file_name, "r") = Null) Then
            fileinput = 1
            fprintf(stderr, "Error opening word file."& vbLf)
            Return -1
        End If
        
        i = 0
        w2memptr = words2mem
        w2offset = 0
        longestlength = 0
        
        While (fgets(buffer, MAX_WORD_LENGTH, word_file_ptr) <> ctype(Null, char))
            j = (CType(strlen(buffer),Integer) - 1)
            buffer(j--) = Microsoft.VisualBasic.ChrW(92)
            strcpy(alphbuffer, alphabetic(buffer))
            If ((CType(strlen(alphbuffer),Integer) < mincwordlength) _
                        OrElse (CType(strlen(alphbuffer),Integer) > maxcwordlength)) Then
                'TODO: Warning!!! continue If
            End If
            
            hasnumber = 0
            j = 0
            Do While (j < CType(strlen(buffer),Integer))
                If ((buffer(j) >= Microsoft.VisualBasic.ChrW(48)) _
                            AndAlso (buffer(j) <= Microsoft.VisualBasic.ChrW(57))) Then
                    hasnumber = 1
                End If
                
                j = (j + 1)
            Loop
            
            If (hasnumber = 1) Then
                'TODO: Warning!!! continue If
            End If
            
            strcpy(ubuffer, uppercase(alphbuffer))
            strcpy(leftover, extract(initword, ubuffer))
            If (leftover(0) = Microsoft.VisualBasic.ChrW(48)) Then
                'TODO: Warning!!! continue If
            End If
            
            strcpy(w2memptr, uppercase(buffer))
            w2memptr = (w2memptr _
                        + (StrLen(buffer) + 1))
            w2offset = (w2offset _
                        + (StrLen(buffer) + 1))
            If (CType(strlen(alphbuffer),Integer) > longestlength) Then
                longestlength = StrLen(alphbuffer)
            End If
            
            If ((w2size - w2offset) _
                        < SAFETY_ZONE) Then
                w2size = (w2size + WORDBLOCKSIZE)
                If (ctype(realloc(words2mem, w2size), char) = ctype(Null, char)) Then
                    fprintf(stderr, "Out of memory.  realloc() returned NULL."& vbLf)
                    exit(-1)
                End If
                
                w2memptr = (words2mem + w2offset)
            End If
            
            i = (i + 1)
            ncount = i
            
        End While
        
        If (fileinput = 1) Then
            fclose (word_file_ptr)
        End If
        
        malloc(ncount, Star, sizeof, (char, Star)
        NULL
        fprintf(stderr, "Insufficient memory.  malloc() returned NULL."& vbLf)
        exit(-1)
        words2(0) = words2mem
        j = 1
        i = 0
        Do While (i < w2size)
            If (j < ncount) Then
                If (words2mem(i) = Microsoft.VisualBasic.ChrW(92)) Then
                    words2(j++) = (words2mem _
                                + (i + 1))
                End If
                
            End If
            
            i = (i + 1)
        Loop
        
        If (silent = 0) Then
            printf(""& vbLf&"%d words loaded (%d byte block).  ", "Longest kept:  %d letters."& vbLf, ncount, w2size, longestlength)
        End If
        
        If (ncount = 0) Then
            If (silent = 0) Then
                printf(""& vbLf&"No candidate words were found, so there are no anagrams."& vbLf)
            End If
            
            exit(0)
        End If
        
        If (CType(malloc(ncount, Star, sizeof, int),Integer) = NULL) Then
            fprintf(stderr, "Insufficient memory; malloc returned NULL."& vbLf)
            exit(-1)
        End If
        
        i = 0
        Do While (i < ncount)
            strcpy(alphbuffer, alphabetic(words2(i)))
            wordsn(i) = CType(strlen(alphbuffer),Integer)
            i = (i + 1)
        Loop
        
        malloc(ncount, Star, sizeof, (char, Star)
        NULL
        fprintf(stderr, "Insufficient memory; malloc returned NULL."& vbLf)
        exit(-1)
        i = 0
        Do While (i < ncount)
            words2ptrs(i) = words2(i)
            i = (i + 1)
        Loop
        
        malloc(ncount, Star, sizeof, (char, Star)
        NULL
        fprintf(stderr, "Insufficient memory; malloc returned NULL."& vbLf)
        exit(-1)
        If (ctype(malloc(w2size, Star, sizeof, char), char) = ctype(Null, char)) Then
            fprintf(stderr, "Insufficient memory; malloc() returned NULL."& vbLf)
            exit(-1)
        End If
        
        keymemptr = keymem
        keyoffset = 0
        i = 0
        Do While (i < ncount)
            strcpy(alphbuffer, alphabetic(words2(i)))
            strcpy(ubuffer, uppercase(alphbuffer))
            strcpy(keymemptr, ubuffer)
            keymemptr = (keymemptr _
                        + (wordsn(i) + 1))
            keyoffset = (keyoffset _
                        + (wordsn(i) + 1))
            i = (i + 1)
        Loop
        
        wordss(0) = keymem
        j = 1
        i = 0
        Do While (i < w2size)
            If (j < ncount) Then
                If (keymem(i) = Microsoft.VisualBasic.ChrW(92)) Then
                    wordss(j++) = (keymem _
                                + (i + 1))
                End If
                
            End If
            
            i = (i + 1)
        Loop
        
        k = 0
        Do While (k < ncount)
            size = CType(strlen(wordss(k)),Integer)
            gap = size
            
            Do Until ((switches <> 0) _
                        Or (gap <> 1))
                gap = max(((gap * 10) _
                                / 13), 1)
                switches = 0
                i = 0
                Do While (i _
                            < (size - gap))
                    j = (i + gap)
                    If (wordss(k)(i) > wordss(k)(j)) Then
                        chold = wordss(k)(i)
                        wordss(k)(i) = wordss(k)(j)
                        wordss(k)(j) = chold
                        switches = (switches + 1)
                    End If
                    
                    i = (i + 1)
                Loop
                
                
            Loop
            
            k = (k + 1)
        Loop
        
        size = ncount
        gap = size
        
        Do Until ((switches <> 0) _
                    Or (gap <> 1))
            gap = max(((gap * 10) _
                            / 13), 1)
            switches = 0
            i = 0
            Do While (i _
                        < (size - gap))
                j = (i + gap)
                If (strcmp(wordss(i), wordss(j)) > 0) Then
                    wholdptr = wordss(i)
                    wordss(i) = wordss(j)
                    wordss(j) = wholdptr
                    wholdptr = words2ptrs(i)
                    words2ptrs(i) = words2ptrs(j)
                    words2ptrs(j) = wholdptr
                    switches = (switches + 1)
                End If
                
                i = (i + 1)
            Loop
            
            
        Loop
        
        largestlet = wordss((ncount - 1))(0)
        size = ncount
        gap = size
        
        Do Until ((switches <> 0) _
                    Or (gap <> 1))
            gap = max(((gap * 10) _
                            / 13), 1)
            switches = 0
            i = 0
            Do While (i _
                        < (size - gap))
                j = (i + gap)
                keyi = wordsn(i)
                keyj = wordsn(j)
                If (keyi > keyj) Then
                    iholdn = wordsn(i)
                    wordsn(i) = wordsn(j)
                    wordsn(j) = iholdn
                    wholdptr = words2(i)
                    words2(i) = words2(j)
                    words2(j) = wholdptr
                    switches = (switches + 1)
                End If
                
                i = (i + 1)
            Loop
            
            
        Loop
        
        If listcandwords Then
            If (silent = 0) Then
                printf(""& vbLf&"List of candidate words:"& vbLf)
            End If
            
            i = 0
            Do While (i < ncount)
                If (silent = 0) Then
                    printf("%6d.  %s"& vbLf, i, words2(i))
                Else
                    printf("%s"& vbLf, words2(i))
                End If
                
                i = (i + 1)
            Loop
            
        End If
        
        If (CType(malloc((longestlength+1Unknown, Star, sizeof, int),Integer) = CType(NULL,Integer)) Then
            fprintf(stderr, "Insufficient memory.  malloc() returned NULL."& vbLf)
            exit(-1)
        End If
        
        If (CType(malloc((longestlength+1Unknown, Star, sizeof, int),Integer) = CType(NULL,Integer)) Then
            fprintf(stderr, "Insufficient memory.  malloc() returned NULL."& vbLf)
            exit(-1)
        End If
        
        i = 0
        Do While (i <= longestlength)
            lindx1(i) = -1
            lindx2(i) = -2
            i = (i + 1)
        Loop
        
        If (ncount > 0) Then
            curpos = 0
            curlen = wordsn(curpos)
            lindx1(curlen) = curpos
            
            Do Until (curpos < ncount)
                
                While (curpos < ncount)
                    If (wordsn(curpos) = curlen) Then
                        curpos = (curpos + 1)
                    Else
                        Exit While
                    End If
                    
                    
                End While
                
                If (curpos >= ncount) Then
                    lindx2(curlen) = (ncount - 1)
                    Exit For
                End If
                
                lindx2(curlen) = (curpos - 1)
                curlen = wordsn(curpos)
                lindx1(curlen) = curpos
                
            Loop
            
        End If
        
        i = 0
        Do While (i < 26)
            findx1(i) = -1
            findx2(i) = -2
            i = (i + 1)
        Loop
        
        If (ncount > 0) Then
            curpos = 0
            curlet = wordss(curpos)(0)
            icurlet = (CType(curlet,Integer) - CType(Microsoft.VisualBasic.ChrW(65),Integer))
            findx1(icurlet) = curpos
            
            Do Until (curpos < ncount)
                
                While (curpos < ncount)
                    If (wordss(curpos)(0) = curlet) Then
                        curpos = (curpos + 1)
                    Else
                        Exit While
                    End If
                    
                    
                End While
                
                If (curpos >= ncount) Then
                    findx2(icurlet) = (ncount - 1)
                    Exit For
                End If
                
                findx2(icurlet) = (curpos - 1)
                curlet = wordss(curpos)(0)
                icurlet = (CType(curlet,Integer) - CType(Microsoft.VisualBasic.ChrW(65),Integer))
                findx1(icurlet) = curpos
                
            Loop
            
        End If
        
        If (CType(malloc(ncount, Star, sizeof, int),Integer) = NULL) Then
            fprintf(stderr, "Insufficient memory; malloc returned NULL."& vbLf)
            exit(-1)
        End If
        
        i = 0
        Do While (i < ncount)
            wordmasks(i) = intmask(wordss(i))
            i = (i + 1)
        Loop
        
        If ((specfirstword = 0) _
                    AndAlso recursiveanag) Then
            If (silent = 0) Then
                printf(""& vbLf&"Anagrams found:"& vbLf)
            End If
            
            malloc(MAX_ANAGRAM_WORDS, Star, sizeof, (char, Star)
            NULL
            fprintf(stderr, "Insufficient memory; malloc returned NULL."& vbLf)
            exit(-1)
            i = 0
            Do While (i < MAX_ANAGRAM_WORDS)
                If (CType(malloc((longestlength+1Unknown, Star, sizeof, char),Char) = CType(NULL,Char)) Then
                    fprintf(stderr, "Insufficient memory; malloc returned NULL."& vbLf)
                    exit(-1)
                End If
                
                i = (i + 1)
            Loop
            
            accum(0)(0) = Microsoft.VisualBasic.ChrW(92)
            level = 0
            rec_anag_count = 0
            minkey = findx1((CType(initword(0),Integer) - CType(Microsoft.VisualBasic.ChrW(65),Integer)))
            anagramr7(initword, accum, minkey, level)
            If (rec_anag_count = 0) Then
                If (silent = 0) Then
                    printf(""& vbLf&"No anagrams found by recursive algorithm."& vbLf)
                End If
                
            End If
            
        End If
        
        If ((specfirstword = 1) _
                    AndAlso recursiveanag) Then
            If (silent = 0) Then
                printf(""& vbLf&"Recursive anagrams found:"& vbLf)
            End If
            
            malloc(MAX_ANAGRAM_WORDS, Star, sizeof, (char, Star)
            NULL
            fprintf(stderr, "Insufficient memory; malloc returned NULL."& vbLf)
            exit(-1)
            i = 0
            Do While (i < MAX_ANAGRAM_WORDS)
                If (CType(malloc((MAX_WORD_LENGTH+1Unknown, Star, sizeof, char),Char) = CType(NULL,Char)) Then
                    fprintf(stderr, "Insufficient memory; malloc returned NULL."& vbLf)
                    exit(-1)
                End If
                
                i = (i + 1)
            Loop
            
            strcpy(accum(0), u_first_word)
            level = 1
            rec_anag_count = 0
            minkey = findx1((CType(remaininitword(0),Integer) - CType(Microsoft.VisualBasic.ChrW(65),Integer)))
            anagramr7(remaininitword, accum, minkey, level)
            If (rec_anag_count = 0) Then
                printf(""& vbLf&"No anagrams found by recursive algorithm."& vbLf)
            End If
            
        End If
        
        Return 0
    End Function
    
    Private Overloads Function uppercase(ByVal s As Char) As Char
        Dim upcasestr() As char
        (MAX_WORD_LENGTH + 1)
        Dim i As Integer
        i = 0
        Do While (i < CType(strlen(s),Integer))
            upcasestr(i) = toupper(s(i))
            i = (i + 1)
        Loop
        
        upcasestr(i) = Microsoft.VisualBasic.ChrW(92)
        Return
    End Function
    
    Private Overloads Function alphabetic(ByVal s As Char) As Char
        Dim alphstr() As char
        (MAX_WORD_LENGTH + 1)
        Dim pos As Integer
        Dim i As Integer
        pos = 0
        i = 0
        Do While (i < CType(strlen(s),Integer))
            If (((s(i) >= Microsoft.VisualBasic.ChrW(65)) _
                        AndAlso (s(i) <= Microsoft.VisualBasic.ChrW(90))) _
                        OrElse ((s(i) >= Microsoft.VisualBasic.ChrW(97)) _
                        AndAlso (s(i) <= Microsoft.VisualBasic.ChrW(122)))) Then
                alphstr(pos++) = s(i)
            End If
            
            i = (i + 1)
        Loop
        
        alphstr(pos) = Microsoft.VisualBasic.ChrW(92)
        Return
    End Function
    
    Private Overloads Function numvowels(ByVal s As Char) As Integer
        Dim vcount As Integer
        Dim cptr As char
        vcount = 0
        cptr = s
        Do While (cptr <> Microsoft.VisualBasic.ChrW(92))
            Select Case (cptr)
                Case Microsoft.VisualBasic.ChrW(65), Microsoft.VisualBasic.ChrW(69), Microsoft.VisualBasic.ChrW(73), Microsoft.VisualBasic.ChrW(79), Microsoft.VisualBasic.ChrW(85), Microsoft.VisualBasic.ChrW(89)
                    vcount = (vcount + 1)
            End Select
            
            cptr = (cptr + 1)
        Loop
        
        Return
    End Function
    
    Private Overloads Sub anagramr7(ByVal s As Char, ByVal Star As Char, ByVal , As accum, ByVal minkey As Integer, ByVal level As Integer)
        Dim s_mask As Integer
        Dim i As Integer
        Dim j As Integer
        Dim extsuccess As Integer
        Dim icurlet As Integer
        Dim newminkey As Integer
        Dim exts() As char
        MAX_WORD_LENGTH
        If (level >= max_depth) Then
            level = (level - 1)
            Return
        End If
        
        If (maxdepthspec = 1) Then
            If (((max_depth - level) _
                        * longestlength) _
                        < CType(strlen(s),Integer)) Then
                level = (level - 1)
                Return
            End If
            
        End If
        
        If (vowelcheck = 1) Then
            If (numvowels(s) = 0) Then
                level = (level - 1)
                Return
            End If
            
        End If
        
        s_mask = intmask(s)
        extsuccess = 0
        icurlet = (CType(s(0),Integer) - CType(Microsoft.VisualBasic.ChrW(65),Integer))
        i = max(minkey, findx1(icurlet))
        Do While (i <= findx2(icurlet))
            If ((s_mask Or wordmasks(i)) _
                        <> s_mask) Then
                'TODO: Warning!!! continue If
            End If
            
            If (adjacentdups = 0) Then
                If ((level > 0) _
                            AndAlso (strcmp(words2ptrs(i), accum((level - 1))) = 0)) Then
                    'TODO: Warning!!! continue If
                End If
                
            End If
            
            strcpy(exts, extract(s, wordss(i)))
            If (exts = Microsoft.VisualBasic.ChrW(48)) Then
                'TODO: Warning!!! continue If
            End If
            
            If (exts = Microsoft.VisualBasic.ChrW(92)) Then
                rec_anag_count = (rec_anag_count + 1)
                strcpy(accum(level), words2ptrs(i))
                If (silent = 0) Then
                    printf("%6d.  ", rec_anag_count)
                End If
                
                j = 0
                Do While (j < level)
                    printf("%s ", accum(j))
                    j = (j + 1)
                Loop
                
                printf("%s"& vbLf, words2ptrs(i))
                extsuccess = 1
                'TODO: Warning!!! continue For
            End If
            
            extsuccess = 1
            strcpy(accum(level), words2ptrs(i))
            level = (level + 1)
            If (adjacentdups = 0) Then
                newminkey = (i + 1)
            Else
                newminkey = i
            End If
            
            anagramr7(exts, accum, newminkey, level)
            i = (i + 1)
        Loop
        
        If (extsuccess = 0) Then
            level = (level - 1)
            Return
        End If
        
        level = (level - 1)
        Return
    End Sub
    
    Private Overloads Function extract(ByVal s1 As Char, ByVal s2 As Char) As Char
        Dim r1() As char
        MAX_WORD_LENGTH
        Dim t1() As char
        MAX_WORD_LENGTH
        Dim Star As char
        Dim s1p As char
        s2p
        r1p
        s1end
        s2end
        Dim s2len As Integer
        Dim found As Integer
        Dim s1len As Integer
        r1p = r1
        strcpy(t1, s1)
        s1p = t1
        s1len = CType(strlen(s1p),Integer)
        s1end = (s1p + s1len)
        s2p = s2
        s2len = CType(strlen(s2),Integer)
        s2end = (s2p + s2len)
        s2p = s2
        Do While (s2p < s2end)
            found = 0
            s1p = t1
            Do While (s1p < s1end)
                If (s2p = s1p) Then
                    s1p = Microsoft.VisualBasic.ChrW(48)
                    found = 1
                    Exit For
                End If
                
                s1p = (s1p + 1)
            Loop
            
            If (found = 0) Then
                r1 = Microsoft.VisualBasic.ChrW(48)
                (r1 + 1) = Microsoft.VisualBasic.ChrW(92)
                Return
            End If
            
            s2p = (s2p + 1)
        Loop
        
        r1p = r1
        s1p = t1
        Do While (s1p < s1end)
            If (s1p <> Microsoft.VisualBasic.ChrW(48)) Then
                (r1p + 1) = s1p
            End If
            
            s1p = (s1p + 1)
        Loop
        
        r1p = Microsoft.VisualBasic.ChrW(92)
        Return
    End Function
    
    Private Overloads Function intmask(ByVal s As Char) As Integer
        Dim sptr As char
        Dim mask As Integer
        mask = 0
        sptr = s
        Do While (sptr <> Microsoft.VisualBasic.ChrW(92))
            If ((sptr >= Microsoft.VisualBasic.ChrW(65)) _
                        AndAlso (sptr <= Microsoft.VisualBasic.ChrW(90))) Then
                mask = (mask Or (1 + CType((sptr - Microsoft.VisualBasic.ChrW(65)),Integer)))
            End If
            
            sptr = (sptr + 1)
        Loop
        
        Return
    End Function

