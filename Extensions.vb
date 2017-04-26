Imports System.Runtime.CompilerServices

Module Extensions

     <Extension()>
     Public Function IsEmpty(ByVal cell As DataGridViewCell) As Boolean
         Return cell.Value Is Nothing OrElse cell.Value Is DBNull.Value OrElse String.IsNullOrWhiteSpace(cell.Value.ToString())
     End Function

     <Extension()>
     Public Function IsNotEmpty(ByVal cell As DataGridViewCell) As Boolean
         Return Not (cell.Value Is Nothing OrElse cell.Value Is DBNull.Value OrElse String.IsNullOrWhiteSpace(cell.Value.ToString()))
     End Function

End Module
