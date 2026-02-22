# ♟️ VBCHESS - Excel VBA Chess Engine By Oily Akara

VBCHESS is a fully playable, feature-rich Chess Engine built entirely inside Microsoft Excel using Visual Basic for Applications (VBA). No external plugins or software are required—just Excel!

Play against a built-in AI, or play against a friend locally with automatic board-flipping and live move notation.

---

## ✨ Features

- **Play vs AI:** Test your skills against a custom chess engine utilizing Alpha-Beta Pruning, Negamax search, and Piece-Square Tables (PST) for positional awareness.
- **Pass & Play (PvP):** Play against a friend locally. The board automatically rotates 180 degrees after every turn so both players get the correct perspective.
- **Modern Graphical UI:** Features a sleek, semi-transparent main menu, animated game-over screens, and a dynamic "Thinking..." turn indicator.
- **Live Match History:** Automatically translates your moves into standard Algebraic Notation (e.g., `Nf3`, `O-O`) and displays them in a clean side panel.
- **Rule Enforcement:** Full legal move generation including En Passant, Castling rights, Check/Checkmate detection, and Draw rules (Threefold Repetition & Insufficient Material).
- **Visual Aids:** Click a piece to see all legal moves highlighted. Kings are highlighted in red when placed in Check.

---

## 🚀 How to Play (Important Installation Steps)

Because this game runs on Excel Macros, Windows will block it by default when you download it. You must unblock it to play.

1. Download the `VBCHESS.xlsm` file to your computer.
2. **Right-click** the downloaded file and select **Properties**.
3. At the bottom of the **General** tab, check the box that says **Unblock**, then click **Apply** and **OK**.
   > If you don't see an Unblock box, you can skip this step.
4. Open the file in **Microsoft Excel**.
5. If Excel prompts a yellow banner at the top saying *"Security Warning: Macros have been disabled"*, click **Enable Content**.
6. The **Main Menu** will appear. Click a mode to start playing!

---

## 🧠 Under the Hood (For Developers)

The engine is written entirely in VBA. Some of the technical highlights include:

- **120-Square Board Representation:** Uses a 10×12 array to easily detect off-board errors during move generation.
- **Move Ordering & Heuristics:** The AI scores captures (MVV-LVA), Killer Moves, and utilizes a History Heuristic to trim the search tree and calculate faster.
- **Quiescence Search:** Prevents the "horizon effect" by continuing to calculate capture sequences even after the target search depth is reached.
- **No API calls:** The game does not connect to Stockfish or any internet databases. The "brain" is 100% contained within the VBA module.

---

## 📸 Screenshots



| Main Menu | In Game |
|:---------:|:-------:|
| ![Main Menu](link_to_image) | ![In Game](link_to_image) |

---

## 🛠️ Compatibility

- Requires **Microsoft Excel for Windows**.
- > **Note:** Excel for Mac does not fully support all VBA Shape rendering features, so UI elements may look slightly different on macOS.