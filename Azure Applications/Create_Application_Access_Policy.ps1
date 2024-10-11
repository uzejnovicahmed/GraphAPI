# Step 1: Connect to Exchange Online
Connect-ExchangeOnline

# Step 2: Create new Service Principal

$APPID = "<APPLICATION ID>"
$ObjectID = "<OBJECT ID>"

# Choose a DisplayName for the Service Principal
$SPDisplayName = "SP_EXO_SEND_MAIL"

# Create your Service Principal
New-ServicePrincipal -AppId $APPID -ObjectId $ObjectID -DisplayName $SPDisplayName

# Get Service Principal information
Get-ServicePrincipal -Identity $ObjectID

# Step 3: Assign Role to the Service Principal
$RoleAssignment = New-ManagementRoleAssignment -App $APPID -Role "Application Mail.Send"

# Display Role Assignment details
$RoleAssignment | Select-Object Name, Role, RoleAssignee, RoleAssigneeType, CustomRecipientWriteScope, AssignmentMethod

# Reference: https://learn.microsoft.com/en-us/graph/auth-limit-mailbox-access

# Step 4: Create a new application access policy to restrict app access

$APPID = "68a677994-ass2-24h4-abc1-a5h35ea1e33c"
$GroupMailAddress = "GRP_LIMIT_ACCESS_APP_EXO_SEND_MAIL@contoso.com"

# Create a new application access policy for the app to restrict access to the APP with the group
New-ApplicationAccessPolicy -AppId $APPID -PolicyScopeGroupId $GroupMailAddress -AccessRight RestrictAccess -Description "Restrict this app to members of group GRP_LIMIT_ACCESS_APP_EXO_SEND_MAIL."

# Step 5: Test the application access policy
Test-ApplicationAccessPolicy -Identity policy.test@contoso.com -AppId $APPID
