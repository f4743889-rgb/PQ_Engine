(raw as nullable text) as text =>
let t = Text.Lower(Text.Trim(raw ?? ""))
in  if Text.EndsWith(t, ".type") then Text.Start(t, Text.Length(t) - 5) else t