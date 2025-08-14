# # List of matches for the 6 teams
# matches_data = [
#     ("Wrexham vs Hull City", [
#         ("Over/Under Goals (Over 2.5)", "Over 2.5", "50%", "-127", "Balanced match—expect some goal threat (Fox Sports)"),
#         ("Over/Under Goals (Under 1.5)", "No", "25%", "+280", "Low likelihood due to attacking styles (Fox Sports)"),
#         ("Both Teams to Score (BTTS)", "Yes", "60%", "-110", "Both sides capable, mid-level scoring (Fox Sports)"),
#         ("Double Chance & Over 1.5 Goals", "Wrexham/Draw + Over 1.5", "65%", "-140", "Home advantage, both teams attacking (Fox Sports)"),
#         ("Double Chance & BTTS", "Wrexham/Draw + BTTS", "55%", "+150", "Safe outcome with likely both scoring (Fox Sports)"),
#         ("Win/Draw/Lose", "Wrexham Win", "45%", "+122", "Home advantage, high momentum (Fox Sports)"),
#         ("Matchwinner & BTTS", "Wrexham & BTTS", "40%", "+250", "Combines win belief with both scoring (Fox Sports)"),
#         ("Fulltime Corners Over/Under (5.5+)", "Over 5.5", "70%", "-200", "Both teams average high corners per game")
#     ]),
#     ("Swansea City vs Crawley Town", [
#         ("Over/Under Goals (Over 2.5)", "Over 2.5", "65%", "-140", "Swansea strong attack, Crawley often concedes"),
#         ("Over/Under Goals (Under 1.5)", "No", "20%", "+300", "Likely multiple goals from Swansea"),
#         ("Both Teams to Score (BTTS)", "Yes", "55%", "+110", "Crawley may score but likely dominated"),
#         ("Double Chance & Over 1.5 Goals", "Swansea/Draw + Over 1.5", "75%", "-180", "Safe pick with goals"),
#         ("Double Chance & BTTS", "Swansea/Draw + BTTS", "50%", "+160", "BTTS less certain"),
#         ("Win/Draw/Lose", "Swansea Win", "70%", "-167", "Favourites by league gap"),
#         ("Matchwinner & BTTS", "Swansea & BTTS", "45%", "+210", "Possible but riskier"),
#         ("Fulltime Corners Over/Under (5.5+)", "Over 5.5", "68%", "-180", "Attacking play likely yields corners")
#     ]),
#     ("Middlesbrough vs Doncaster Rovers", [
#         ("Over/Under Goals (Over 2.5)", "Over 2.5", "60%", "-120", "Boro at home usually score multiple"),
#         ("Over/Under Goals (Under 1.5)", "No", "25%", "+280", "Lower division opponents may concede often"),
#         ("Both Teams to Score (BTTS)", "No", "55%", "+100", "Doncaster likely struggle to score"),
#         ("Double Chance & Over 1.5 Goals", "Middlesbrough/Draw + Over 1.5", "70%", "-160", "Safe outcome + goals"),
#         ("Double Chance & BTTS", "Middlesbrough/Draw + BTTS", "45%", "+200", "Lower BTTS likelihood"),
#         ("Win/Draw/Lose", "Middlesbrough Win", "65%", "-150", "Favourites but cup risk exists"),
#         ("Matchwinner & BTTS", "Middlesbrough & BTTS", "40%", "+230", "Possible but less likely"),
#         ("Fulltime Corners Over/Under (5.5+)", "Over 5.5", "72%", "-190", "Strong attack drives corners")
#     ]),
#     ("Stockport County vs Crewe Alexandra", [
#         ("Over/Under Goals (Over 2.5)", "Over 2.5", "62%", "-125", "Both teams high goal rate recently"),
#         ("Over/Under Goals (Under 1.5)", "No", "28%", "+260", "Chances of low-scoring match slim"),
#         ("Both Teams to Score (BTTS)", "Yes", "58%", "-110", "Crewe often score even when losing"),
#         ("Double Chance & Over 1.5 Goals", "Stockport/Draw + Over 1.5", "65%", "-140", "Safe with likely goals"),
#         ("Double Chance & BTTS", "Stockport/Draw + BTTS", "55%", "+150", "Both scoring probable"),
#         ("Win/Draw/Lose", "Stockport Win", "55%", "-123", "Favourites but close match"),
#         ("Matchwinner & BTTS", "Stockport & BTTS", "45%", "+210", "Likely if both score"),
#         ("Fulltime Corners Over/Under (5.5+)", "Over 5.5", "70%", "-180", "Attacking styles yield corners")
#     ]),
#     ("Grimsby Town vs Shrewsbury Town", [
#         ("Over/Under Goals (Over 2.5)", "Over 2.5", "50%", "+100", "Evenly matched sides, goals possible"),
#         ("Over/Under Goals (Under 1.5)", "Yes", "30%", "+240", "Could be tight defensively"),
#         ("Both Teams to Score (BTTS)", "Yes", "52%", "+105", "Similar strengths, likely exchange goals"),
#         ("Double Chance & Over 1.5 Goals", "Grimsby/Draw + Over 1.5", "60%", "-125", "Safer pick with goals"),
#         ("Double Chance & BTTS", "Grimsby/Draw + BTTS", "55%", "+140", "Both teams may score"),
#         ("Win/Draw/Lose", "Draw", "35%", "+220", "Even match potential stalemate"),
#         ("Matchwinner & BTTS", "Grimsby & BTTS", "40%", "+250", "Possible home edge"),
#         ("Fulltime Corners Over/Under (5.5+)", "Over 5.5", "65%", "-160", "Even matches often high corners")
#     ]),
#     ("Coventry City vs Luton Town", [
#         ("Over/Under Goals (Over 2.5)", "Over 2.5", "55%", "-110", "Both in decent scoring form"),
#         ("Over/Under Goals (Under 1.5)", "No", "25%", "+280", "Low scoring unlikely"),
#         ("Both Teams to Score (BTTS)", "Yes", "58%", "-110", "Both teams strong attacking options"),
#         ("Double Chance & Over 1.5 Goals", "Coventry/Draw + Over 1.5", "63%", "-135", "Safe + goals"),
#         ("Double Chance & BTTS", "Coventry/Draw + BTTS", "55%", "+140", "Likely both teams score"),
#         ("Win/Draw/Lose", "Coventry Win", "50%", "+100", "Slight home edge"),
#         ("Matchwinner & BTTS", "Coventry & BTTS", "45%", "+210", "Balanced bet"),
#         ("Fulltime Corners Over/Under (5.5+)", "Over 5.5", "68%", "-170", "Both sides win many corners")
#     ])
# ]

# # Flatten data for DataFrame
# full_data = []
# for match, bets in matches_data:
#     for market, prediction, probability, odds, rationale in bets:
#         full_data.append([match, market, prediction, probability, odds, rationale])

# # Create DataFrame
# df_all = pd.DataFrame(full_data, columns=["Match", "Market", "Prediction", "Probability", "Odds", "Rationale/Notes"])

# # Save to Excel
# file_path_all = "/mnt/data/EFL_Cup_Multi_Market_Predictions.xlsx"
# df_all.to_excel(file_path_all, index=False)

# file_path_all
