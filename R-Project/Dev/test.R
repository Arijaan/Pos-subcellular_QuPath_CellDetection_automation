library(svDialogs)

message <- "Please enter your name for the report:"
default_value <- Sys.info()["user"] # Uses the system username as a default

user_response <- dlg_input(
    message = message,
    default = default_value
)$res