VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Day13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Days")
Option Explicit
    '========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C

    ' This is the opposite of minimum route problem.  Here we are trying to find the maximum route (maximum happiness)
    ' the cooplication here is the the relationship between the two nodes is not symmetrical
    ' i.e. A ->B has a different value to B->A
    Private Const TODAY             As String = "\Day13.txt"

    Private Type State

        Data                            As seqC
        Guests                          As KvpC
        GuestPermutations               As seqC
    End Type

    Private s                           As State


    Public Sub Execute()
        
            Part01
            Part02
            
    End Sub

    Private Sub Part01()

        Initialise
        maths.Permutations s.Guests.Keys, s.Guests.Count, s.GuestPermutations
        
        ' s.GuestPermutations now contains all possible unique arrangements of guests
        ' For each permutations we need to examine the happiness level between two guests by
        ' looking up the gursts in the s.guests kvps.
        Dim myMaxhappiness As Long: myMaxhappiness = globals.MinLong
        
        Dim myGuestPermutations As Iteritems: Set myGuestPermutations = Iteritems(s.GuestPermutations)
        Do
            Dim myGuestPermutation As seqC: Set myGuestPermutation = myGuestPermutations.curitem(0)
            Dim myHappiness As Long: myHappiness = 0
            Dim myGuests As Iteritems: Set myGuests = Iteritems.Deb(myGuestPermutation)
            ' we are iterating a single guest plan
            Do
                Dim myGuest1 As String
                Dim myGuest2 As String
                If myGuests.HasNoNext Then

                    myGuest1 = myGuests.curitem(0)
                    myGuest2 = myGuests.curitem(-(myGuestPermutation.Count - 1))


                Else

                    myGuest1 = myGuests.curitem(0)
                    myGuest2 = myGuests.curitem(1)


                End If
                
                myHappiness = myHappiness + s.Guests.Item(myGuest1).Item(myGuest2)
                myHappiness = myHappiness + s.Guests.Item(myGuest2).Item(myGuest1)
                
            Loop While myGuests.MoveNext
            
            If myHappiness > myMaxhappiness Then
                myMaxhappiness = myHappiness
            End If
          
            
        Loop While myGuestPermutations.MoveNext
        
        Dim myResult As Long: myResult = myMaxhappiness
        fmt.dbg "The answer to Day {0} Part 01 is 709.  Found is {1}", VBA.Mid$(TODAY, 5, 2), myResult
        
    End Sub

    Private Sub Part02()
        ' The number of guests is increased by 1 by adding the host
        ' the happines score for sitting next to a host is always 0
        ' thus the solution for part 2 uses the same code as for part 1 with the
        ' exception that we don't need to check for the case of wrap around
        
        Initialise
        
        maths.Permutations s.Guests.Keys, s.Guests.Count, s.GuestPermutations
        
        Dim myMaxhappiness As Long: myMaxhappiness = globals.MinLong
        
        Dim myGuestPermutations As Iteritems: Set myGuestPermutations = Iteritems(s.GuestPermutations)
        Do
            Dim myGuestPermutation As seqC: Set myGuestPermutation = myGuestPermutations.curitem(0)
            Dim myHappiness As Long: myHappiness = 0
            Dim myGuests As Iteritems: Set myGuests = Iteritems.Deb(myGuestPermutation)
            ' we are iterating a single guest plan
            
            Do While myGuests.MoveNext
                Dim myGuest1 As String
                Dim myGuest2 As String

                myGuest1 = myGuests.curitem(-1)
                myGuest2 = myGuests.curitem(0)

                myHappiness = myHappiness + s.Guests.Item(myGuest1).Item(myGuest2)
                myHappiness = myHappiness + s.Guests.Item(myGuest2).Item(myGuest1)
                
            Loop
            
            If myHappiness > myMaxhappiness Then
                myMaxhappiness = myHappiness
            End If
          
            
        Loop While myGuestPermutations.MoveNext
        
        Dim myResult As Long: myResult = myMaxhappiness
        
        fmt.dbg "The answer to Day {0} part 2 is 668.  Found is {1}", VBA.Mid$(TODAY, 5, 2), myResult

    End Sub
    
    

    ' Process the incoming data to a suitable format
    Private Sub Initialise()
    
        ' The input strings are of the form
        'Alice would gain 54 happiness units by sitting next to Bob.
        'or
        'Alice would lose 54 happiness units by sitting next to Bob.
        ' we replace 'lose ' with a minus sign to get a negative number, e.g. lose 54 become -54
        ' we delete the extraneous Would, gain , happiness units by sitting next to and the .
        ' we replace multiple spaces by a single space to give 'Alice 54 Bob'
        ' and finally we split the string at the spaces into a seq
        
        Set s.Data = seqC(Filer.GetFileAsArrayOfStrings(AoCRawData & Year & TODAY, vbCrLf)) _
            .MapIt(mpMultiReplace(seqC(Array("gain ", vbNullString), Array("lose ", "-"), Array("happiness units by sitting next to", vbNullString), Array("would ", vbNullString), Array(chars.twPeriod, vbNullString), Array(chars.twSpace & chars.twSpace, chars.twSpace)))) _
            .MapIt(mpsplit(chars.twSpace))
        ' s.data comprises a seq of guest1, happiness, score guest 2
        ' Create a dictionary of guest 1 vs dictionary of guest 2 happiness score
        ' we now have the abiity to get the happiness score of any pairing
        ' guest1 vs guest 2 and the reverse guest 2 vs guest1 so we gcan calculate the total happiness
        ' of any two guests sitting next to each other
        Set s.Guests = KvpC.Deb
            
        Dim myGuests As Iteritems: Set myGuests = Iteritems.Deb(s.Data)
        Do
            Dim myGuest As seqC: Set myGuest = myGuests.curitem(0)
            Dim myLHS As String: myLHS = myGuest.Item(1)
            Dim myHappiness As Long: myHappiness = VBA.CLng(myGuest.Item(2))
            Dim myRHS As String: myRHS = myGuest.Item(3)
            If s.Guests.HoldsKey(myLHS) Then
                s.Guests.Item(myLHS).Add myRHS, myHappiness
            Else
                s.Guests.Add myLHS, KvpC.Deb.Add(myRHS, myHappiness)
            End If
            
        Loop While myGuests.MoveNext
        
        ' plans will be a sequence of permutations of the guests
        Set s.GuestPermutations = seqC.Deb
        
    End Sub
    
    

