/* global Word, Office, document */

const TNR = "Times New Roman";
const BN_FONT = "Kalpurush";
const EN_DASH = "\u2013";
const PRIME = "\u2032";
const MID_DOT = "\u00B7";
const TIMES = "\u00D7";
const INTEGRAL = "\u222B";
const CHUNK = 40; // কতগুলো ফিল্ড একসাথে প্রসেস হবে

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("run").onclick = run;
  }
});

function log(msg) {
  const el = document.getElementById("status");
  el.textContent += msg + "\n";
}

// Word এ "Replace All" নেই, তাই সব মিল খুঁজে একে একে বদলাই
async function replaceAll(context, scope, find, replacement, matchCase = false) {
  const hits = scope.search(find, { matchCase });
  hits.load("items");
  await context.sync();
  hits.items.forEach((r) => r.insertText(replacement, "Replace"));
  await context.sync();
  return hits.items.length;
}

// \s\up0(x) বা \s\do0(x) -> x (সুপারস্ক্রিপ্ট/সাবস্ক্রিপ্ট), EQ এর বাইরে
async function scriptify(context, body, tag, prop) {
  const pattern = "\\\\s\\\\" + tag + "0\\([!\\(\\)]@\\)";
  const hits = body.search(pattern, { matchWildcards: true });
  hits.load("items/text");
  await context.sync();
  hits.items.forEach((r) => {
    const m = r.text.match(/\(([\s\S]*)\)$/);
    const nr = r.insertText(m ? m[1] : "", "Replace");
    nr.font[prop] = true;
  });
  await context.sync();
  return hits.items.length;
}

async function run() {
  const btn = document.getElementById("run");
  const lOnlyEq = document.getElementById("lEq").checked;
  btn.disabled = true;
  document.getElementById("status").textContent = "";

  try {
    await Word.run(async (context) => {
      const body = context.document.body;

      // STEP 1: হাইফেন -> en dash, ' ও ’ -> prime, \dot ও \cdot -> middle dot
      log("Step 1: চিহ্ন বদলানো...");
      await replaceAll(context, body, "-", EN_DASH);
      await replaceAll(context, body, "'", PRIME);
      await replaceAll(context, body, "\u2019", PRIME);
      await replaceAll(context, body, "\\dot", MID_DOT);
      await replaceAll(context, body, "\\cdot", MID_DOT);

      // STEP 2: ফন্ট ও সাইজ
      log("Step 2: ফন্ট সেট করা...");
      body.font.name = TNR;
      body.font.size = 10.5;
      body.font.bold = false;
      body.font.italic = false;
      await context.sync();

      // বাংলা অক্ষরে Kalpurush
      const bn = body.search("[\u0980-\u09FF\u0964\u0965]{1,}", { matchWildcards: true });
      bn.load("items");
      await context.sync();
      bn.items.forEach((r) => (r.font.name = BN_FONT));
      await context.sync();

      // STEP 2.5 ও 3: {eq ...} -> EQ ফিল্ড
      log("Step 3: EQ ফিল্ড বানানো...");
      await replaceAll(context, body, "{eq ", "{ EQ ", false);
      const found = body.search("\\{ EQ [!\\}]@\\}", { matchWildcards: true });
      found.load("items/text");
      await context.sync();
      for (let i = found.items.length - 1; i >= 0; i--) {
        const r = found.items[i];
        const inner = r.text.slice(1, -1).trim().replace(/^EQ\s*/i, "");
        r.insertField("Replace", Word.FieldType.eq || "Eq", inner, false);
      }
      await context.sync();
      log("  " + found.items.length + " টি EQ ফিল্ড তৈরি হয়েছে");

      // এখন সব EQ ফিল্ড লোড করি
      const fields = body.fields;
      fields.load("items/type");
      await context.sync();
      const eqs = fields.items.filter((f) => f.type === "Eq");
      log("  ডকুমেন্টে মোট EQ ফিল্ড: " + eqs.length);

      // STEP 4, 5: ফিল্ডের ভেতরের ফরম্যাটিং
      log("Step 4-5: ফিল্ডের ভেতরের ফরম্যাটিং...");
      const fieldOps = [
        { find: EN_DASH, apply: (r) => (r.font.name = TNR) },
        { find: TIMES, apply: (r) => (r.font.name = TNR) },
        { find: INTEGRAL, apply: (r) => (r.font.name = TNR) },
      ];
      ["4", "7", "10"].forEach((n) =>
        fieldOps.push({
          find: "\\\\s\\\\up" + n + "\\([!\\)]@\\)",
          wild: true,
          apply: (r) => (r.font.size = 7.5),
        })
      );
      ["2", "8"].forEach((n) =>
        fieldOps.push({
          find: "\\\\s\\\\do" + n + "\\([!\\)]@\\)",
          wild: true,
          apply: (r) => (r.font.size = 7.5),
        })
      );
      if (lOnlyEq) {
        fieldOps.push({ find: "l", matchCase: true, apply: (r) => (r.font.italic = true) });
        ["lim", "log", "ln"].forEach((fn) =>
          fieldOps.push({ find: fn, matchCase: true, apply: (r) => (r.font.italic = false) })
        );
      }

      for (let s = 0; s < eqs.length; s += CHUNK) {
        const part = eqs.slice(s, s + CHUNK);

        // সাধারণ অপারেশন
        const jobs = [];
        part.forEach((f) =>
          fieldOps.forEach((op) => {
            const c = f.code.search(op.find, {
              matchWildcards: !!op.wild,
              matchCase: !!op.matchCase,
            });
            c.load("items");
            jobs.push({ c, op });
          })
        );
        await context.sync();
        jobs.forEach(({ c, op }) => c.items.forEach(op.apply));
        await context.sync();

        // STEP 5.5: \s(...) এর ভেতরে ∫ থাকলে লিমিট 7.5, ∫ নিজে 13.5
        const blocks = part.map((f) => {
          const c = f.code.search("\\\\s\\([!\\)]@\\)", { matchWildcards: true });
          c.load("items/text");
          return c;
        });
        await context.sync();
        const ints = [];
        blocks.forEach((c) =>
          c.items.forEach((r) => {
            if (r.text.indexOf(INTEGRAL) !== -1) {
              r.font.size = 7.5;
              const i = r.search(INTEGRAL);
              i.load("items");
              ints.push(i);
            }
          })
        );
        await context.sync();
        ints.forEach((i) =>
          i.items.forEach((r) => {
            r.font.size = 13.5;
            r.font.name = TNR;
          })
        );
        await context.sync();
      }

      // EQ এর বাইরের সাধারণ টেক্সটে একই ফন্ট ও 'l' ফরম্যাটিং
      for (const ch of [EN_DASH, TIMES, INTEGRAL]) {
        const h = body.search(ch);
        h.load("items");
        await context.sync();
        h.items.forEach((r) => (r.font.name = TNR));
        await context.sync();
      }
      if (!lOnlyEq) {
        const ls = body.search("l", { matchCase: true });
        ls.load("items");
        await context.sync();
        ls.items.forEach((r) => (r.font.italic = true));
        await context.sync();
        for (const fn of ["lim", "log", "ln"]) {
          const h = body.search(fn, { matchCase: true });
          h.load("items");
          await context.sync();
          h.items.forEach((r) => (r.font.italic = false));
          await context.sync();
        }
      }

      // STEP 6 ও 7: \s\up0( ) ও \s\do0( ) -> সুপার/সাবস্ক্রিপ্ট
      log("Step 6-7: সুপারস্ক্রিপ্ট/সাবস্ক্রিপ্ট...");
      const nUp = await scriptify(context, body, "up", "superscript");
      const nDo = await scriptify(context, body, "do", "subscript");
      log("  up: " + nUp + ", do: " + nDo);

      // STEP 8: শুধু EQ ফিল্ড আপডেট ও লক
      log("Step 8: ফিল্ড আপডেট ও লক...");
      eqs.forEach((f) => {
        try {
          f.updateResult();
        } catch (e) {
          /* আপডেট না হলেও চলবে */
        }
      });
      await context.sync();
      eqs.forEach((f) => (f.locked = true));
      await context.sync();
    });
    log("সম্পন্ন।");
  } catch (e) {
    log("সমস্যা: " + e.message);
    console.error(e);
  } finally {
    btn.disabled = false;
  }
}
