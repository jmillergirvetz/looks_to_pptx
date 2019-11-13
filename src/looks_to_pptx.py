from looker_sdk import client, models, error

# client calls will now automatically authenticate using the
# api3credentials specified in 'looker.ini'
sdk = client.setup("looker.ini")
looker_api_user = sdk.me()