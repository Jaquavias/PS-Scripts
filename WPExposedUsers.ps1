$Target = "https://www.domain.com/wp-json/wp/v2/users"

try {
    $response = Invoke-WebRequest -Uri $Target -UseBasicParsing
    $users = $response.Content | ConvertFrom-Json

    foreach ($user in $users) {
        Write-Host "ID: $($user.id) | Name: $($user.name) | Slug: $($user.slug)"
    }
} catch{
    Write-Host "Failed to get users loser"
}