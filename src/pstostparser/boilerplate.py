"""Boilerplate stripping for email bodies.

Removes machine-added notice blocks (external-sender warnings, confidentiality
footers, AV scan lines, ...) so converted Markdown contains the human content.

Custom patterns: pass a text file via --boilerplate-file. Each non-blank,
non-# line is either a literal substring (case-insensitive) or a regex if
wrapped in slashes, e.g.::

    CAUTION: external sender
    /Do not (click|open)[^.]{0,200}unexpected[^.]{0,200}\\./
"""

from __future__ import annotations

import re

# (label, regex) — order matters; specific first, generic last.
DEFAULT_PATTERNS: list[tuple[str, str]] = [
    ("external-org-warning",
     r"[ \t]*This email comes from outside the organization\.?"
     r"[\s\S]*?unless it is an email you expected to receive\.?"),
    ("external-beware",
     r"BEWARE\s*:?\s*This\s+email\s+originated\s+outside\s+of\s+our\s+organization\.?"
     r"[\s\S]{0,300}?DO\s+NOT\s+CLICK[\s\S]{0,300}?content\s+is\s+safe\.?"),
    ("warning-external-email",
     r"WARNING:\s*External\s+email\.\s*Do\s+not\s+click\s+or\s+open\s+attachments"
     r"\s+unless\s+you\s+know\s+and\s+trust\s+the\s+sender\.?"),
    ("info-are-confidential",
     r"The\s+information\s+in\s+this\s+email\s+and\s+in\s+any\s+attachments\s+are"
     r"\s+confidential[\s\S]{0,600}?distribute\s+any\s+part\s+of\s+its\s+contents\.?"),
    ("note-ebridge",
     r"Note:\s*This\s+message\s+may\s+contain\s+confidential\s+information\.?"
     r"[\s\S]{0,900}?independent\s+of\s+this\s+message\.?"),
    ("confidential-contained-nor-disclose",
     r"Confidential\s+information\s+may\s+be\s+contained\s+in\s+this\s+message\.?"
     r"[\s\S]{0,400}?nor\s+disclose\s+its\s+contents\s+to\s+any\s+other\s+person\.?"),
    ("do-not-disseminate-part",
     r"Please\s+do\s+not\s+copy,\s+distribute,\s+or\s+disseminate\s+part\s+or\s+whole"
     r"\s+of\s+this\s+email[\s\S]{0,600}?to\s+and\s+from\s+our\s+systems\.?"),
    ("print-refrain",
     r"please\s+refrain\s+from\s+printing\s+emails\.?"),
    ("external-ttj-bracket",
     r"\[This email originated from outside of [^\]\n]{1,60}\]"),
    ("external-originated-standalone",
     r"[ \t]*(\[EXTERNAL( EMAIL)?\]\s*:?\s*)?"
     r"This (email|message) originated from outside\s+(of\s+)?"
     r"(the\s+organization|our\s+organization|[A-Z][\w&.']{1,40})\.?\]?"),
    ("external-caution",
     r"(CAUTION|WARNING)\s*:?\s*This (email|message) (originated|came|was sent)"
     r" from outside[\s\S]{0,400}?"
     r"(do not (click|open)[^.\n]*\.|verify the sender[^.\n]*\.)"),
    ("external-caution-is-external",
     r"(C[Aa]+UTION|ATTENTION|WARNING)\s*:?\s*This is an EXTERNAL (email|e-?mail)"
     r"[\s\S]{0,400}?DO NOT (open|click)[\s\S]{0,300}?"
     r"(unknown senders|unexpected (email|message))[^\n.]{0,150}\.?"),
    ("external-do-not-click",
     r"Do not click links or open attachments[^.\n]{0,200}\."),
    ("confidentiality-disclaimer",
     r"This (e-?mail|message) (and any (files|attachments)[\s\S]{0,120}?)?"
     r"(is |are )?(confidential|privileged|intended solely)[\s\S]{0,800}?"
     r"(notify the sender|delete (it|this (e-?mail|message)))[^\n.]{0,150}\.?"),
    ("confidential-destined-only",
     r"This email and attachments are confidential and destined only"
     r" for\s+the\s+use\s+of the intended\s+addressees\.?"
     r"[\s\S]{0,200}?notify\s+the\s+sender\.?"),
    ("files-transmitted-confidential",
     r"This email and any files transmitted with it are confidential"
     r"[\s\S]{0,600}?please\s+contact\s+the\s+sender\.?"),
    ("sender-communication-confidential",
     r"The information contained in this communication from the sender"
     r" is confidential\.?[\s\S]{0,600}?strictly\s+prohibited"
     r"(\s+and\s+may\s+be\s+unlawful)?\.?"),
    ("confidential-contents-info",
     r"The\s+contents\s+and\s+information\s+in\s+this\s+email\s+including"
     r"\s+attachments\s+are\s+confidential"
     r"[\s\S]{0,900}?delete\s+this\s+email\s+entirely\.?"),
    ("confidential-may-be-in-message",
     r"(Privileged/\s*)?Confidential\s+information\s+may\s+be\s+(in|contained\s+in)"
     r"\s+this\s+message\.?[\s\S]{0,300}?delete\s+the\s+email"
     r"(\s+and\s+notify\s+the\s+sender\s+immediately\s+if\s+you\s+receive\s+this\s+in\s+error)?\.?"),
    ("delete-email-notify-sender",
     r"Please\s+delete\s+the\s+email\s+and\s+notify\s+the\s+sender\s+immediately"
     r"\s+if\s+you\s+receive\s+this\s+in\s+error\.?"),
    ("privileged-may-be-contained",
     r"Privileged/Confidential information may\s+be\s+contained\s+in\s+this\s+message\.?"),
    ("official-secrets-act",
     r"Communication of any information in this email to any unauthorized\s+person"
     r"\s+is\s+an offence[\s\S]{0,400}?receive\s+this\s+in\s+error\.?\"?"),
    ("message-contains-confidential",
     r"This\s+message\s+contains\s+confidential\s+information"
     r"\s+and\s+is\s+intended\s+only\s+for\s+the\s+individual\s+named\.?"
     r"[\s\S]{0,600}?delete\s+this\s+email\s+from\s+your\s+system\.?"),
    ("recipient-of-communication",
     r"If\s+you\s+are\s+not\s+the\s+intended\s+recipient\s+of\s+this\s+communication,"
     r"\s+please\s+notify\s+the\s+sender\s+and\s+delete\s+it\.?"
     r"[\s\S]{0,500}?Official\s+Secrets\s+Act\s*\(Cap\s+213\)"),
    ("caution-external-party",
     r"CAUTION\s*:?\s*This\s+email\s+originated\s+from\s+an\s+external\s+party"
     r"\s+outside[\s\S]{0,200}?unless\s+you\s+know\s+the\s+sender\.?"),
    ("dextra-privileged",
     r"Privileged,\s+confidential\s+and/or\s+copyright\s+information\s+may\s+be"
     r"\s+contained\s+in\s+this\s+e-?mail\.?"
     r"[\s\S]{0,800}?delete\s+this\s+e-?mail\s+from\s+your\s+system\.?"),
    ("integrity-not-guaranteed",
     r"As\s+the\s+integrity\s+of\s+this\s+message\s+cannot\s+be\s+guaranteed"
     r"[\s\S]{0,900}?may\s+not\s+necessarily\s+represent\s+the\s+opinion"
     r"[^\n.]{0,150}\.?"),
    ("classification-confidential",
     r"(?m)^\s*Classification:\s*Confidential\s*$"),
    ("bare-confidential-line",
     r"(?m)^\s*Confidential\s*$"),
    ("bare-privileged-prefix",
     r"(?m)^Privileged/\s*$"),
    ("notify-sender-standalone",
     r"(Privileged/\s*)?Please\s+notify\s+the\s+sender\s+immediately\s+if\s+you\s+receive\s+this\s+in\s+error\.?"),
    ("save-paper-think-green",
     r"Save\s+paper,\s+think\s+green\s+before\s+you\s+print\.?"),
    ("our-environment-print",
     r"Please\s+consider\s+our\s+environment"
     r"\s+before\s+printing[^\n.]{0,120}\.?"),
    ("transmission-intended-addressee",
     r"This e-?mail transmission is intended only for the addressee\.?"
     r"[\s\S]{0,400}?erase all copies of the message and attachment\(s\)\.?"),
    ("notice-privileged-communication",
     r"NOTICE\s*:\s*This communication and any attachments"
     r"[\s\S]{0,1000}?destroy any printed copies\.?"),
    ("confidentiality-notice-wsp",
     r"CONFIDENTIALITY\s+NOTICE:\s*This\s+email,\s+including\s+any\s+attachments,"
     r"\s+is\s+confidential[\s\S]{0,1000}?monitor\s+all\s+email\s+communications"
     r"\s+through\s+its\s+networks\.?"),
    ("named-addressee-only",
     r"This email is intended to be received only by the named addressee\.?"
     r"[\s\S]{0,900}?delete\s+the\s+email\s+from\s+all\s+your\s+records\.?"),
    ("no-liability-virus",
     r"We\s+do\s+not\s+accept\s+liability\s+in\s+connection\s+with\s+computer\s+virus"
     r"[\s\S]{0,300}?unauthori[sz]ed amendment\.?"),
    ("transferred-intended-only",
     r"The\s+information\s+transferred\s+is\s+intended\s+only\s+for\s+the\s+person"
     r"\s+or\s+entity\s+to\s+which\s+it\s+is\s+addressed"
     r"\s+and\s+may\s+contain\s+confidential\s+and/or\s+privileged\s+material\.?"
     r"[\s\S]{0,400}?not\s+intended\s+to\s+waive\s+confidentiality"
     r"\s+and/or\s+privilege\.?"),
    ("subject-to-contract",
     r"The\s+contents\s+of\s+this\s+email\s+are\s+subject\s+to\s+contract"
     r"[\s\S]{0,700}?terms-and-conditions"),
    ("do-not-disclose-rely",
     r"If\s+you\s+are\s+not\s+the\s+intended\s+recipient\s+please\s+do\s+not"
     r"\s+disclose,\s+copy,\s+distribute\s+or\s+take\s+any\s+action\s+in\s+reliance"
     r"\s+on\s+the\s+contents\s+of\s+this\s+e-?mail\.?"),
    ("sgs-block",
     r"Information\s+in\s+this\s+email\s+and\s+any\s+attachments\s+is\s+confidential"
     r"[\s\S]{0,900}?accessible\s+at\s+https://www\.sgs\.com/en/terms-and-conditions"),
    ("bentley-virus-free",
     r"Bentley\s+Systems\s+has\s+taken\s+all\s+reasonable\s+steps"
     r"[\s\S]{0,400}?as\s+a\s+result\s+of\s+this\s+or\s+any\s+communication\s+by\s+email\.?"),
    ("conf-caution-including",
     r"CONFIDENTIALITY\s+CAUTION:\s*\(including\s+the\s+attachments\)\.?"
     r"[\s\S]{0,900}?unauthorised\s+use\s+by\s+any\s+other\s+person\.?"),
    ("note-message-confidential",
     r"NOTE:\s*message,\s+including\s+any\s+attachments,\s+contains\s+confidential"
     r"\s+information[\s\S]{0,500}?strictly\s+prohibited\.?"),
    ("dangling-if-you-are-not",
     r"(?m)^If\s+you\s+are\s+not\s*$"),
    ("bracketed-confidentiality",
     r"\[CONFIDENTIALITY:\s*If\s+you\s+receive\s+this\s+email\s+in\s+error"
     r"[\s\S]{0,800}?criminal\s+offence\.\]"),
    ("caution-sent-from-outside",
     r"CAUTION:\s*This\s+message\s+was\s+sent\s+from\s+outside\s+of\s+"
     r"[\w&.,'\-]+\.?[\s\S]{0,300}?report\s+all\s+suspicious\s+emails\s+to\s+\S+\.?"),
    ("warning-external-sender",
     r"\*{0,3}\s*WARNING:\s*this\s+message\s+is\s+from\s+an\s+EXTERNAL\s+SENDER\.?"
     r"\s*Please\s+be\s+cautious,\s+particularly\s+with\s+links\s+and\s+attachments\.?"
     r"\s*\*{0,3}"),
    ("jenmon-all-information",
     r"All\s+information\s+in\s+this\s+email\s+is\s+intended\s+for\s+the\s+use"
     r"\s+of\s+the\s+addressee\s+solely[\s\S]{0,2000}?"
     r"check\s+this\s+email\s+and\s+attachments\s+for\s+viruses\.?"),
    ("teams-unmonitored-mailbox",
     r"This\s+email\s+was\s+sent\s+from\s+an\s+unmonitored\s+mailbox\.?"
     r"[\s\S]{0,400}?Redmond\s+WA\s+98052-7329"
     r"(\s*Read\s+our\s+privacy\s+policy\.?)?"),
    ("may-contain-confidential",
     r"This\s+email,\s+including\s+any\s+attachments,\s+may\s+contain\s+confidential"
     r"\s+and/or[\s\S]{0,600}?delete\s+this\s+message\s+and\s+any\s+copies\.?"),
    ("disclaimer-email-privacy",
     r"Disclaimer & Email Privacy Statement\s*:?"
     r"[\s\S]{0,1200}?Thank\s+you\s+for\s+your\s+co-operation\.?"),
    ("bracketed-confidential",
     r"\[This email \(including any attachments\) is confidential"
     r" and may also be privileged\.?[\s\S]{0,1000}?Thank you\.\]"),
    ("unauthorised-disclosure-reliance",
     r"Any\s+unauthorised\s+disclosure,\s+copying,\s+distribution\s+and/or"
     r"\s+taking\s+of\s+any\s+action\s+in\s+reliance[\s\S]{0,200}?"
     r"strictly\s+prohibited\.?(\s+and\s+may\s+be\s+unlawful\.?)?"),
    ("creeg-disclaimer",
     r"Disclaimer:\s+The\s+above\s+e-?mail\s+message,\s+its\s+attachments"
     r"\s+and\s+included\s+replies[\s\S]{0,1200}?"
     r"Thank\s+you\s+for\s+your\s+cooperation\.?"),
    ("comfortdelgro-statements",
     r"Statements\s+pertaining\s+to\s+any\s+matter\s+outside\s+our\s+business"
     r"\s+are\s+not\s+to\s+be\s+taken\s+as\s+endorsed[\s\S]{0,700}?"
     r"signed\s+by\s+the\s+duly\s+authori[sz]ed\s+representatives"
     r"\s+of\s+both\s+parties\.?"),
    ("print-only-if-necessary",
     r"We\s+encourage\s+you\s+to\s+print\s+this\s+only\s+if\s+necessary\.?"),
    ("warning-not-addressee",
     r'WARNING\s*:?\s*"?\s*If\s+you\s+are\s+not\s+the\s+intended\s+addressee,\s+you\s+must\s+not\s+copy,\s+distribute\s+or\s+take\s+any\s+action\s+in\s+reliance\s+thereon\.?'),
    ("not-recipient-do-not-copy",
     r"If\s+you\s+are\s+not\s+the\s+intended\s+recipient,\s+you\s+must\s+not\s+copy,"
     r"\s+distribute\s+or\s+use\s+it\s+for\s+any\s+purpose"
     r"[\s\S]{0,400}?notify\s+the\s+sender\s+immediately"
     r"\s+if\s+you\s+receive\s+this\s+in\s+error\.?"),
    ("bare-caution-header",
     r"(?m)^\s*CAUTION\s*:?\s*$"),
    ("bare-caution-external",
     r"(?m)^\s*CAUTION:\s*EXTERNAL\s+EMAIL\s*$"),
    ("warning-cyber-fraud",
     r"WARNING\s+-\s+To\s+minimi[sz]e\s+the\s+risk\s+of\s+cyber\s+fraud"
     r"[\s\S]{0,600}?correct\s+bank\s+account\s+details\.?"
     r"(\s+If\s+you\s+receive\s+an\s+email\s+purporting[\s\S]{0,400}?by\s+phone\.?)?"),
    ("confidentiality-caution-msg",
     r"CONFIDENTIALITY\s+CAUTION\s*:?\s*This\s+message\s+is\s+intended\s+only"
     r"\s+for\s+the\s+use\s+of\s+the\s+individual\s+or\s+entity"
     r"[\s\S]{0,900}?delete\s+the\s+original\s+message\.?"),
    ("named-addressee-plural",
     r"This\s+e-?mail\s+is\s+intended\s+only\s+for\s+the\s+named\s+addressee\(s\)"
     r"[\s\S]{0,600}?strictly\s+prohibited\.?"),
    ("av-footer",
     r"This (e-?mail|message) has been (checked|scanned)[^\n]*viruses?[^\n]*|"
     r"Virus-?free\.?[^\n]*(avast|avg|mcafee|symantec|kaspersky)[^\n]*"),
    ("sent-from-mobile",
     r"(?m)^Sent from my (iPhone|iPad|Android|Samsung|Windows Phone|mobile( device)?)\s*$"),
    ("think-before-print",
     r"Please\s+consider\s+the\s+environment"
     r"[^\n]*befor[^\n]*printing[^\n.]{0,80}\.?"),
    ("bare-think-green",
     r"(?m)^Think\s+Green\s*:?\s*$"),
    ("unsubscribe",
     r"If you no longer wish to receive[^\n]*unsubscribe[^\n]*\.?|"
     r"To unsubscribe[^\n]{0,200}\.?"),
]

_COMPILED: list[tuple[str, re.Pattern]] = [
    (label, re.compile(pat, re.IGNORECASE)) for label, pat in DEFAULT_PATTERNS
]


def load_custom_patterns(path: str) -> list[tuple[str, re.Pattern]]:
    """Parse a custom boilerplate file into (label, compiled-regex) pairs."""
    out: list[tuple[str, re.Pattern]] = []
    with open(path, encoding="utf-8-sig") as f:
        for i, raw in enumerate(f, 1):
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            label = f"custom-{i}"
            if len(line) >= 2 and line.startswith("/") and line.rfind("/") > 0:
                pat = line[1:line.rfind("/")]
            else:
                pat = re.escape(line)
            try:
                out.append((label, re.compile(pat, re.IGNORECASE)))
            except re.error as e:
                raise ValueError(f"{path}:{i}: bad regex {line!r}: {e}")
    return out


def strip_boilerplate(
    text: str,
    extra: list[tuple[str, re.Pattern]] | None = None,
) -> tuple[str, list[str]]:
    """Return (cleaned_text, [labels of removed blocks])."""
    if not text:
        return text, []
    removed: list[str] = []
    for label, rx in list(_COMPILED) + list(extra or []):
        text, n = rx.subn("", text)
        if n:
            removed.append(label if n == 1 else f"{label}x{n}")
    text = re.sub(r"\n{3,}", "\n\n", text).strip()
    return text, removed
