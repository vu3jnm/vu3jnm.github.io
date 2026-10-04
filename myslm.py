#!/usr/bin/env python3
"""
tiny_lm.py - Train and sample from a small character-level GPT on any text file.

Install:   pip install torch
Train:     python tiny_lm.py train --data mytext.txt --steps 3000
Generate:  python tiny_lm.py generate --prompt "Once upon" --max_new_tokens 400

Tips:
  * Works best with at least ~100 KB of text (1 MB+ is better).
  * On CPU, keep the defaults (small model). On a GPU, raise --n_layer,
    --n_embd and --block_size for noticeably better output.
"""
import argparse
import math
import os

import torch
import torch.nn as nn
import torch.nn.functional as F


# --------------------------------------------------------------------------
# Tokenizer (character level)
# --------------------------------------------------------------------------
class CharTokenizer:
    def __init__(self, chars):
        self.chars = list(chars)
        self.stoi = {c: i for i, c in enumerate(self.chars)}

    @classmethod
    def from_text(cls, text):
        return cls(sorted(set(text)))

    @property
    def vocab_size(self):
        return len(self.chars)

    def encode(self, s):
        # Characters unseen during training are skipped.
        return [self.stoi[c] for c in s if c in self.stoi]

    def decode(self, ids):
        return "".join(self.chars[i] for i in ids)


# --------------------------------------------------------------------------
# Model
# --------------------------------------------------------------------------
class CausalSelfAttention(nn.Module):
    def __init__(self, n_embd, n_head, dropout):
        super().__init__()
        assert n_embd % n_head == 0, "n_embd must be divisible by n_head"
        self.n_head = n_head
        self.qkv = nn.Linear(n_embd, 3 * n_embd)
        self.proj = nn.Linear(n_embd, n_embd)
        self.dropout = dropout
        self.resid_drop = nn.Dropout(dropout)

    def forward(self, x):
        B, T, C = x.shape
        q, k, v = self.qkv(x).split(C, dim=2)
        q, k, v = (t.view(B, T, self.n_head, C // self.n_head).transpose(1, 2)
                   for t in (q, k, v))
        y = F.scaled_dot_product_attention(
            q, k, v, is_causal=True,
            dropout_p=self.dropout if self.training else 0.0)
        y = y.transpose(1, 2).contiguous().view(B, T, C)
        return self.resid_drop(self.proj(y))


class Block(nn.Module):
    def __init__(self, n_embd, n_head, dropout):
        super().__init__()
        self.ln1 = nn.LayerNorm(n_embd)
        self.attn = CausalSelfAttention(n_embd, n_head, dropout)
        self.ln2 = nn.LayerNorm(n_embd)
        self.mlp = nn.Sequential(
            nn.Linear(n_embd, 4 * n_embd),
            nn.GELU(),
            nn.Linear(4 * n_embd, n_embd),
            nn.Dropout(dropout),
        )

    def forward(self, x):
        x = x + self.attn(self.ln1(x))
        return x + self.mlp(self.ln2(x))


class TinyGPT(nn.Module):
    def __init__(self, vocab_size, block_size, n_layer, n_head, n_embd, dropout):
        super().__init__()
        self.block_size = block_size
        self.tok_emb = nn.Embedding(vocab_size, n_embd)
        self.pos_emb = nn.Embedding(block_size, n_embd)
        self.drop = nn.Dropout(dropout)
        self.blocks = nn.Sequential(
            *[Block(n_embd, n_head, dropout) for _ in range(n_layer)])
        self.ln_f = nn.LayerNorm(n_embd)
        self.head = nn.Linear(n_embd, vocab_size, bias=False)
        self.head.weight = self.tok_emb.weight  # weight tying
        self.apply(self._init)

    @staticmethod
    def _init(m):
        if isinstance(m, nn.Linear):
            nn.init.normal_(m.weight, mean=0.0, std=0.02)
            if m.bias is not None:
                nn.init.zeros_(m.bias)
        elif isinstance(m, nn.Embedding):
            nn.init.normal_(m.weight, mean=0.0, std=0.02)

    def forward(self, idx, targets=None):
        B, T = idx.shape
        pos = torch.arange(T, device=idx.device)
        x = self.drop(self.tok_emb(idx) + self.pos_emb(pos))
        x = self.ln_f(self.blocks(x))
        logits = self.head(x)
        loss = None
        if targets is not None:
            loss = F.cross_entropy(logits.view(-1, logits.size(-1)),
                                   targets.view(-1))
        return logits, loss

    @torch.no_grad()
    def generate(self, idx, max_new_tokens, temperature=0.8, top_k=40):
        self.eval()
        for _ in range(max_new_tokens):
            idx_cond = idx[:, -self.block_size:]
            logits, _ = self(idx_cond)
            logits = logits[:, -1, :] / max(temperature, 1e-5)
            if top_k is not None:
                v, _ = torch.topk(logits, min(top_k, logits.size(-1)))
                logits[logits < v[:, [-1]]] = -float("inf")
            probs = F.softmax(logits, dim=-1)
            idx = torch.cat([idx, torch.multinomial(probs, 1)], dim=1)
        return idx


# --------------------------------------------------------------------------
# Training
# --------------------------------------------------------------------------
def get_batch(data, block_size, batch_size, device):
    ix = torch.randint(len(data) - block_size - 1, (batch_size,))
    x = torch.stack([data[i:i + block_size] for i in ix])
    y = torch.stack([data[i + 1:i + block_size + 1] for i in ix])
    return x.to(device), y.to(device)


@torch.no_grad()
def estimate_loss(model, splits, args, device, iters=20):
    model.eval()
    out = {}
    for name, data in splits.items():
        losses = torch.zeros(iters)
        for k in range(iters):
            x, y = get_batch(data, args.block_size, args.batch_size, device)
            losses[k] = model(x, y)[1].item()
        out[name] = losses.mean().item()
    model.train()
    return out


def lr_at(step, args):
    """Linear warmup, then cosine decay to 10% of the peak learning rate."""
    if step < args.warmup:
        return args.lr * (step + 1) / args.warmup
    progress = (step - args.warmup) / max(1, args.steps - args.warmup)
    return args.lr * (0.1 + 0.9 * 0.5 * (1 + math.cos(math.pi * progress)))


def pick_device():
    if torch.cuda.is_available():
        return "cuda"
    if getattr(torch.backends, "mps", None) and torch.backends.mps.is_available():
        return "mps"
    return "cpu"


def train(args):
    device = pick_device()
    torch.manual_seed(args.seed)

    with open(args.data, "r", encoding="utf-8") as f:
        text = f.read()
    tok = CharTokenizer.from_text(text)
    data = torch.tensor(tok.encode(text), dtype=torch.long)

    n_val = max(args.block_size + 2, int(0.1 * len(data)))
    if len(data) < 2 * (args.block_size + 2):
        raise SystemExit(
            f"Text too short ({len(data)} chars) for block_size={args.block_size}. "
            "Provide more text or lower --block_size.")
    splits = {"train": data[:-n_val], "val": data[-n_val:]}

    cfg = dict(vocab_size=tok.vocab_size, block_size=args.block_size,
               n_layer=args.n_layer, n_head=args.n_head,
               n_embd=args.n_embd, dropout=args.dropout)
    model = TinyGPT(**cfg).to(device)
    n_params = sum(p.numel() for p in model.parameters())
    print(f"device={device} | chars={len(text):,} | vocab={tok.vocab_size} "
          f"| params={n_params / 1e6:.2f}M")

    opt = torch.optim.AdamW(model.parameters(), lr=args.lr,
                            betas=(0.9, 0.95), weight_decay=0.1)

    def save():
        torch.save({"config": cfg, "chars": tok.chars,
                    "state_dict": model.state_dict()}, args.out)

    best_val = float("inf")
    for step in range(args.steps + 1):
        for g in opt.param_groups:
            g["lr"] = lr_at(step, args)

        if step % args.eval_every == 0 or step == args.steps:
            losses = estimate_loss(model, splits, args, device)
            print(f"step {step:5d} | train {losses['train']:.4f} "
                  f"| val {losses['val']:.4f}")
            if losses["val"] < best_val:  # keep the best checkpoint
                best_val = losses["val"]
                save()

        if step == args.steps:
            break

        x, y = get_batch(splits["train"], args.block_size,
                         args.batch_size, device)
        _, loss = model(x, y)
        opt.zero_grad(set_to_none=True)
        loss.backward()
        torch.nn.utils.clip_grad_norm_(model.parameters(), 1.0)
        opt.step()

    print(f"Done. Best val loss {best_val:.4f}. Saved to {args.out}")
    sample(args, quiet=True)


# --------------------------------------------------------------------------
# Generation
# --------------------------------------------------------------------------
def sample(args, quiet=False):
    device = pick_device()
    ckpt = torch.load(args.out, map_location=device)
    tok = CharTokenizer(ckpt["chars"])
    model = TinyGPT(**ckpt["config"]).to(device)
    model.load_state_dict(ckpt["state_dict"])

    prompt = getattr(args, "prompt", "") or ""
    ids = tok.encode(prompt) or [0]
    idx = torch.tensor([ids], dtype=torch.long, device=device)
    out = model.generate(idx, args.max_new_tokens,
                         temperature=args.temperature, top_k=args.top_k)
    if not quiet:
        print("-" * 60)
    print(tok.decode(out[0].tolist()))


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------
def main():
    p = argparse.ArgumentParser(description="Tiny character-level GPT")
    sub = p.add_subparsers(dest="cmd", required=True)

    t = sub.add_parser("train", help="train on a text file")
    t.add_argument("--data", required=True, help="path to a UTF-8 text file")
    t.add_argument("--out", default="tiny_lm.pt")
    t.add_argument("--steps", type=int, default=3000)
    t.add_argument("--batch_size", type=int, default=32)
    t.add_argument("--block_size", type=int, default=128, help="context length")
    t.add_argument("--n_layer", type=int, default=4)
    t.add_argument("--n_head", type=int, default=4)
    t.add_argument("--n_embd", type=int, default=128)
    t.add_argument("--dropout", type=float, default=0.1)
    t.add_argument("--lr", type=float, default=3e-3)
    t.add_argument("--warmup", type=int, default=100)
    t.add_argument("--eval_every", type=int, default=250)
    t.add_argument("--seed", type=int, default=1337)
    t.add_argument("--max_new_tokens", type=int, default=300)
    t.add_argument("--temperature", type=float, default=0.8)
    t.add_argument("--top_k", type=int, default=40)
    t.add_argument("--prompt", default="")

    g = sub.add_parser("generate", help="sample from a trained model")
    g.add_argument("--out", default="tiny_lm.pt", help="checkpoint path")
    g.add_argument("--prompt", default="")
    g.add_argument("--max_new_tokens", type=int, default=400)
    g.add_argument("--temperature", type=float, default=0.8)
    g.add_argument("--top_k", type=int, default=40)

    args = p.parse_args()
    train(args) if args.cmd == "train" else sample(args)


if __name__ == "__main__":
    main()

# pip install torch
# python myslm.py train --data mytext.txt --steps 3000
# python myslm.py generate --prompt "Once upon" --max_new_tokens 400