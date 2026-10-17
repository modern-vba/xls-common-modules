Attribute VB_Name = "Lib_IPv4"
Option Explicit
Option Base 0
Option Private Module

' #############################################################################
'!
'! @brief
'! Standard module that groups IPv4-related functions and helpers.
'! It also contains members not used by this tool because they are shared with other tools.
'!
' #############################################################################

Private Const C_OCTET_SEP_RE As String = "\."
Private Const C_OCTET_SEP As String = "."
Private Const C_MASK_SEP As String = "/"
Private Const C_MASK_SEP_RE As String = "[/_]"

'* Regular expression representing one octet of an IPv4 address string.
'*
'* @note Decimal notation only.
Public Const G_IPV4_OCTET_RE As String = "(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])"

'* Regular expression representing one octet of an IPv4 mask string.
'*
'* @note Decimal notation only.
Public Const G_IPV4_MOCTET_RE As String = "(0|128|192|224|240|248|252|254|255)"

'* Regular expression representing the mask length of an IPv4 address string.
Public Const G_IPV4_MASKLEN_RE As String = "(3[0-2]|[1-2][0-9]|[0-9])"

'* Regular expression representing the entire IPv4 address string (ex. 192.168.1.2).
'*
'* @details
'* The structure is "(1st)\.(2nd)\.(3rd)\.(4th)". Use this as a reference when using submatches.
Public Const G_IPV4_ADDR_RE As String = G_IPV4_OCTET_RE & C_OCTET_SEP_RE & G_IPV4_OCTET_RE & C_OCTET_SEP_RE & G_IPV4_OCTET_RE & C_OCTET_SEP_RE & G_IPV4_OCTET_RE

'* Regular expression representing the entire IPv4 mask string (ex. 255.255.255.0).
'*
'* @details
'* The structure is "((1st is 0-254, rest are 0)|(1st is 255, 2nd is 0-254, rest are 0)|(255 continues and 3rd is 0-254, rest are 0)|(255 continues and last is 0-255))".
'* Use this as a reference when using submatches.
Public Const G_IPV4_MASK_RE As String = "(" & _
        "(0|128|192|224|240|248|252|254)" & C_OCTET_SEP_RE & "(0)" & C_OCTET_SEP_RE & "(0)" & C_OCTET_SEP_RE & "(0)" & "|" & _
        "(255)" & C_OCTET_SEP_RE & "(0|128|192|224|240|248|252|254)" & C_OCTET_SEP_RE & "(0)" & C_OCTET_SEP_RE & "(0)" & "|" & _
        "(255)" & C_OCTET_SEP_RE & "(255)" & C_OCTET_SEP_RE & "(0|128|192|224|240|248|252|254)" & C_OCTET_SEP_RE & "(0)" & "|" & _
        "(255)" & C_OCTET_SEP_RE & "(255)" & C_OCTET_SEP_RE & "(255)" & C_OCTET_SEP_RE & "(0|128|192|224|240|248|252|254|255)" & _
        ")"

'* Regular expression representing an IPv4 network address string (ex. 192.168.1.0/255.255.255.0 or 192.168.1.0/24).
'*
'* @details
'* ÅuG_IPV4_ADDR_RE & "(|" & C_MASK_SEP_RE & G_IPV4_MASK_RE & "|" & C_MASK_SEP_RE & G_IPV4_MASKLEN_RE & ")"Åv
'* Use this as a reference when using submatches.
Public Const G_IPV4_NW_RE As String = G_IPV4_ADDR_RE & "(|" & C_MASK_SEP_RE & G_IPV4_MASK_RE & "|" & C_MASK_SEP_RE & G_IPV4_MASKLEN_RE & ")"

'* Gets the broader network address (mask length shorter by 1).
'*
'* @param NetworkAddressValue Network address value.
'* @param MaskLength Current mask length.
'* @return Broader network address value.
'*
'* @details
'* Calculates a network address whose mask length is shorter by 1 based on the specified network address and mask length.
'* Raises an error when the mask length is 0 or the specified address is not a network address.
Public Function ExpandNetwork(ByVal NetworkAddressValue As Long, ByVal MaskLength As Integer) As Long
    Dim result_value As Long
    Dim mask_value As Long

    If MaskLength = 0 Then
        Err.Raise Number:=vbObjectError + 1, Source:="Function ExpandNetwork", Description:="Cannot expand the network any further. (" & ConvertToIpAddress(NetworkAddressValue) & "/" & MaskLength & ")"
        Exit Function
    End If

    mask_value = ConvertFromMaskLength(MaskLength)
    If Not IsNetwork(NetworkAddressValue, mask_value) Then
        Err.Raise Number:=vbObjectError + 1, Source:="Function ExpandNetwork", Description:="The value is not a network address. (" & ConvertToIpAddress(NetworkAddressValue) & "/" & MaskLength & ")"
        Exit Function
    End If

    mask_value = ConvertFromMaskLength(MaskLength - 1)
    result_value = GetNetworkAddress(NetworkAddressValue, mask_value)

    ExpandNetwork = result_value
End Function

'* Gets narrower network addresses (subnets with mask length longer by 1).
'*
'* @param NetworkAddressValue Network address value.
'* @param MaskLength Current mask length.
'* @return Array of two subnet addresses.
'*
'* @details
'* Calculates two subnet addresses whose mask length is longer by 1 based on the specified network address and mask length.
'* Raises an error when the mask length has reached 32 or the specified address is not a network address.
Public Function NarrowNetwork(ByVal NetworkAddressValue As Long, ByVal MaskLength As Integer) As Long()
    Dim result_value() As Long
    ReDim result_value(0 To 1) As Long
    Dim mask_value As Long

    If 29 < MaskLength And MaskLength < 33 Then
        Err.Raise Number:=vbObjectError + 1, Source:="Function NarrowNetwork", Description:="Cannot narrow the network any further. (" & ConvertToIpAddress(NetworkAddressValue) & "/" & MaskLength & ")"
        Exit Function
    End If

    mask_value = ConvertFromMaskLength(MaskLength)
    If Not IsNetwork(NetworkAddressValue, mask_value) Then
        Err.Raise Number:=vbObjectError + 1, Source:="Function NarrowNetwork", Description:="The value is not a network address. (" & ConvertToIpAddress(NetworkAddressValue) & "/" & MaskLength & ")"
        Exit Function
    End If

    result_value(0) = NetworkAddressValue
    result_value(1) = NetworkAddressValue Or GetMaskValue(MaskLength)

    NarrowNetwork = result_value
End Function

'* Formats an IP address or network address into the appropriate form.
'*
'* @param AddressLike String to format. IP address format is assumed.
'* @return Formatted address string.
'*
'* @details
'* This function replaces specific characters (underscores) in the input address string with slashes,
'* and formats it as standard CIDR notation. If the input address does not include a network mask,
'* "/32" is appended to the end, producing a form that represents an individual IP address.
'*
'* Specifically, it replaces "_" in the input string with "/" and, when the input form is a single IP address,
'* appends "/32" to format it in CIDR notation. If a network mask is already specified,
'* returns the form as is.
Public Function WellFormedAddress(ByVal AddressLike As String) As String
    Dim result_value As String

    result_value = Replace(AddressLike, "_", "/")

    If result_value <> "" And result_value Like "?*.?*.?*.?*" And Not result_value Like "?*.?*.?*.?*/?*" Then
        result_value = result_value & "/32"
    End If

    WellFormedAddress = result_value
End Function

'* Parses a string in "IP address/mask" format and converts it to values.
'*
'* @param IpAddressValue [Output] Parsed IP address value.
'* @param MaskValue [Output] Parsed mask value.
'* @param MaskLength [Output] Parsed mask length.
'* @param IpAddressAndMask String in "IP address/mask" format to parse.
'*
'* @details
'* Parses a string in "IP address/mask" format and converts it to IP address value, mask value, and mask length.
'* The mask part accepts both CIDR notation and dotted-decimal format.
'* Raises an error if a string with an invalid format is passed.
Public Sub ParseIpAddressAndMask(ByRef IpAddressValue As Long, ByRef MaskValue As Long, ByRef MaskLength As Integer, ByVal IpAddressAndMask As String)
    Dim ip_arr() As String
    Dim ip_addr_value As Long
    Dim mask_value As Long
    Dim mask_length As Integer

    ip_arr = Split(IpAddressAndMask, C_MASK_SEP)

    If 1 < UBound(ip_arr) Then
        Err.Raise Number:=vbObjectError + 1, Source:="Sub ParseIpAddressAndMask", Description:="The format is invalid. (" & IpAddressAndMask & ")"
    End If

    ip_addr_value = ConvertFromIpAddress(ip_arr(0))

    'If IsIpAddress(ip_arr(1)) Then
    If Not IsInteger(ip_arr(1)) Then
        mask_value = ConvertFromIpAddress(ip_arr(1))
        If Not IsValidMaskValue(mask_value) Then
            Err.Raise Number:=vbObjectError + 1, Source:="Sub ParseIpAddressAndMask", Description:="The format is invalid. (" & IpAddressAndMask & ")"
            Exit Sub
        End If
        mask_length = ConvertToMaskLength(mask_value)
    Else
        mask_length = CInt(ip_arr(1))
        mask_value = ConvertFromMaskLength(mask_length)
    End If

    IpAddressValue = ip_addr_value
    MaskValue = mask_value
    MaskLength = mask_length
End Sub

'* Checks whether the given IP address value is a network address.
'*
'* @param IpAddressValue IP address value to check.
'* @param MaskValue Mask value to use for checking.
'* @return True if it is a network address; otherwise, False.
'*
'* @details
'* Checks whether it is a network address based on the specified IP address value and mask value.
Public Function IsNetwork(ByVal IpAddressValue As Long, ByVal MaskValue As Long) As Boolean
    Dim result_value As Boolean
    Dim host_addr As Long

    host_addr = GetHostAddress(IpAddressValue, MaskValue)

    If host_addr = 0 Then
        IsNetwork = True
    Else
        IsNetwork = False
    End If
End Function

''
'' Checks whether the value satisfies the IP address format (dotted decimal notation).
''
'Function IsIpAddress(IpAddress As String) As Long
'    Dim regex_obj As New RegExp
'    Dim match_collection As MatchCollection
'
'    regex_obj.Pattern = "^" & IPV4_OCTET_RE & OCTET_SEP_RE & IPV4_OCTET_RE & OCTET_SEP_RE & IPV4_OCTET_RE & OCTET_SEP_RE & IPV4_OCTET_RE & "$"
'    regex_obj.Global = False
'
'    Set match_collection = regex_obj.Execute(IpAddress)
'    If 0 < match_collection.count Then
'        IsIpAddress = True
'    Else
'        IsIpAddress = False
'    End If
'End Function

'* Checks whether the mask consists of contiguous 1 bits from the beginning.
'*
'* @param MaskValue Mask value to check.
'* @return True if the mask consists of contiguous 1 bits from the beginning; otherwise, False.
'*
'* @details
'* Checks whether the specified mask value is a valid mask consisting of contiguous 1 bits from the beginning.
'* Returns False for an invalid mask (when it contains 0 in the middle).
Public Function IsValidMaskValue(ByVal MaskValue As Long) As Boolean
    Dim found_zero As Boolean
    Dim item_idx As Integer

    If MaskValue > 0 Then
        ' Positive numbers have a leading 0, followed by at least one 1.
        IsValidMaskValue = False
        Exit Function
    End If

    If MaskValue = 0 Then
        ' Return a fixed value because the digit would overflow.
        ' All zeros.
        IsValidMaskValue = True
        Exit Function
    End If

    ' Check all 32 bits.
    For item_idx = 30 To 0 Step -1
        If (MaskValue And (2& ^ item_idx)) <> 0 Then
            ' If the bit is set.
            If found_zero Then
                ' If a 0 has already been found.
                IsValidMaskValue = False
                Exit Function
            End If
        Else
            ' The first 0 was found.
            found_zero = True
        End If
    Next item_idx

    ' All ones.
    IsValidMaskValue = True
End Function

'* Gets the host part of an IP address.
'*
'* @param IpAddressValue Target IP address value.
'* @param MaskValue Netmask value to use.
'* @return Value indicating the host part.
'*
'* @details
'* Applies the netmask value to the specified IP address value and extracts the host part.
Public Function GetHostAddress(ByVal IpAddressValue As Long, ByVal MaskValue As Long) As Long
    GetHostAddress = IpAddressValue And Not MaskValue
End Function

'* Gets the network part of an IP address.
'*
'* @param IpAddressValue Target IP address value.
'* @param MaskValue Netmask value to use.
'* @return Value indicating the network part.
'*
'* @details
'* Applies the netmask value to the specified IP address value and extracts the network part.
Public Function GetNetworkAddress(ByVal IpAddressValue As Long, ByVal MaskValue As Long) As Long
    GetNetworkAddress = IpAddressValue And MaskValue
End Function

'* Converts an IP address string to an IP address value.
'*
'* @param IpAddress IP address string to convert.
'* @return Corresponding IP address value.
'*
'* @details
'* Converts a dotted-decimal IP address string to a Long IP address value.
'* Raises an error for invalid formats.
Public Function ConvertFromIpAddress(ByVal IpAddress As String) As Long
    Dim regex_obj As RegExp
    Dim match_collection As MatchCollection
    Dim match_item As Match
    Dim item_idx As Integer
    Dim ipv4_re_obj As String
    Dim octet_value As Long
    Dim high_bit As Long
    Dim result_value As Long

    Set regex_obj = New RegExp

    regex_obj.Pattern = "^" & G_IPV4_ADDR_RE & "$"
    regex_obj.Global = False

    Set match_collection = regex_obj.Execute(IpAddress)

    If 0 < match_collection.Count Then
        Set match_item = match_collection(0)
        If match_item.SubMatches.Count = 4 Then
            octet_value = CInt(match_item.SubMatches(0))

            ' Prevent overflow.
            If octet_value < 128 Then
                high_bit = 0
            Else
                octet_value = &H7F And octet_value
                high_bit = &H80000000
            End If

            result_value = high_bit Or (octet_value * 16777216 + CLng(match_item.SubMatches(1)) * 65536 + CLng(match_item.SubMatches(2)) * 256 + CLng(match_item.SubMatches(3)))

            ' Exit after successful completion.
            ConvertFromIpAddress = result_value
            Exit Function
        End If
    End If

    ' Raise an error if completion was not successful.
    Err.Raise Number:=vbObjectError + 1, Source:="Function ConvertFromIpAddress", Description:="The address format is invalid. (" & IpAddress & ")"
End Function

'* Converts an IP address value to an IP address string.
'*
'* @param IpAddressValue IP address value to convert.
'* @return Corresponding IP address string.
'*
'* @details
'* Converts the specified Long IP address value to a dotted-decimal string.
Public Function ConvertToIpAddress(ByVal IpAddressValue As Long) As String
    Dim result_value As String
    Dim ip_value As Long
    Dim mask_value As Long
    Dim high_bit As Long

    ' Create a mask that extracts only the lower 8 bits.
    mask_value = GetMaskValue(24, 32)
    'mask_value = BitRight(&HFFFFFFFF, 24)

    ' Handle only the most significant bit separately.
    If 0 <= IpAddressValue Then
        ip_value = IpAddressValue
        high_bit = 0
    Else
        ip_value = &H7FFFFFFF And IpAddressValue
        high_bit = &H80
    End If

    ' Fourth octet.
    result_value = ip_value And mask_value
    ip_value = ip_value \ 256

    ' Third octet.
    result_value = "" & (ip_value And mask_value) & C_OCTET_SEP & result_value
    ip_value = ip_value \ 256

    ' Second octet.
    result_value = "" & (ip_value And mask_value) & C_OCTET_SEP & result_value
    ip_value = ip_value \ 256

    ' First octet.
    result_value = "" & (high_bit Or ip_value) & C_OCTET_SEP & result_value

    ConvertToIpAddress = result_value
End Function

'* Inverts a mask value.
'*
'* @param MaskValue Mask value to invert.
'* @return Inverted mask value.
'*
'* @details
'* Inverts the specified mask value bit by bit.
Public Function InvertMaskValue(ByVal MaskValue As Long) As Long
    InvertMaskValue = Not MaskValue
End Function

'* Converts a mask length to a mask value.
'*
'* @param MaskLength Mask length to convert (within the range 0 to 32).
'* @return Corresponding mask value.
'*
'* @details
'* Converts the specified mask length to a 32-bit mask value.
'* Raises an error for invalid mask lengths (out of range).
Public Function ConvertFromMaskLength(ByVal MaskLength As Integer) As Long
    If MaskLength < 0 Or 32 < MaskLength Then
        Err.Raise Number:=vbObjectError + 1, Source:="Function ConvertFromMaskLength", Description:="Mask length is out of range. (" & MaskLength & ")"
        Exit Function
    End If

    ConvertFromMaskLength = GetMaskValue(0, MaskLength)
    'ConvertFromMaskLength = BitLeft(&HFFFFFFFF, (32 - MaskLength))
End Function

'* Converts a mask value to a mask length.
'*
'* @param MaskValue Mask value to convert.
'* @return Corresponding mask length.
'*
'* @details
'* Converts the specified mask value to a mask length.
'* Raises an error for invalid mask values (when they are not contiguous 1 bits).
Public Function ConvertToMaskLength(ByVal MaskValue As Long) As Integer
    Dim found_zero As Boolean
    Dim mask_length As Integer
    Dim item_idx As Integer

    If MaskValue > 0 Then
        Err.Raise vbObjectError + 1, Source:="Function ConvertToMaskLength", Description:="Invalid mask value. (" & MaskValue & ")"
        Exit Function
    End If

    If MaskValue = 0 Then
        ' Return a fixed value because the digit would overflow.
        ConvertToMaskLength = 0
        Exit Function
    End If

    ' Check all 32 bits.
    mask_length = 1
    For item_idx = 30 To 0 Step -1
        If (MaskValue And (2& ^ item_idx)) <> 0 Then
            If found_zero Then
                Err.Raise vbObjectError + 1, Source:="ConvertToMaskLength", Description:="Invalid mask value. (" & MaskValue & ")"
            End If
            mask_length = mask_length + 1
        Else
            found_zero = True
        End If
    Next item_idx

    ConvertToMaskLength = mask_length
End Function

'* Creates a mask with the specified bit range set to 1.
'*
'* @param Start Start bit (0-origin, inclusive).
'* @param Finish Finish bit (0-origin, exclusive). If omitted, Start + 1.
'* @return Created mask value.
'*
'* @details
'* Creates a 32-bit mask value with bits from the start bit through the finish bit set to 1 and all other bits set to 0.
'* Raises an error for out-of-range bit specifications.
Public Function GetMaskValue(ByVal Start As Integer, Optional ByVal Finish As Integer = -1) As Long
    Dim result_value As Long
    Dim start_idx As Long
    Dim high_bit As Long
    Dim bit_length As Integer

    If Finish = -1 Then
        Finish = Start + 1
    End If

    If Finish < Start Or Start < 0 Or 32 < Finish Then
        Err.Raise vbObjectError + 1, Source:="Function GenerateMask", Description:="The mask start and finish positions are invalid. (" & Start & ", " & Finish & ")"
    End If


    ' Process the most significant bit.
    If Start = 0 And Start <> Finish Then
        high_bit = &H80000000
        start_idx = 1
    Else
        high_bit = 0
        start_idx = Start
    End If

    bit_length = Finish - start_idx

    result_value = (2& ^ bit_length - 1) * (2& ^ (32 - Finish)) Or high_bit

    GetMaskValue = result_value
End Function
