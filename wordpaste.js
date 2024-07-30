function copyToClipboard() {
    navigator.clipboard.writeText(me('#pasteArea').textContent).then(_ => {
        let btn = me('#copyBtn');
        btn.textContent = 'Copied!';
        sleep(2000).then(_ => btn.textContent = 'Copy');
    });
}

// Remainder of this file is from OpenSlides/openslides-client with minor changes

function unwrapNode(node) {
    const parent = node.parentNode;
    while (node.firstChild) parent.insertBefore(node.firstChild, node);
    parent.removeChild(node);
}

function parseRomanNumber(roman) {
    roman = roman.toUpperCase();
    let value = 0;
    const values = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };
    let i = roman.length;
    let lastVal = 0;
    while (i--) {
        if (values[roman.charAt(i)] >= lastVal) value += values[roman.charAt(i)];
        else value -= values[roman.charAt(i)];
        lastVal = values[roman.charAt(i)];
    }
    return value;
}

function parseLetterNumber(str) {
    const alphaVal = (s) => s.toLowerCase().charCodeAt(0) - 97 + 1;
    let value = 0;
    let i = str.length;
    while (i--) {
        const factor = Math.pow(26, str.length - i - 1);
        value += alphaVal(str.charAt(i)) * factor;
    }
    return value;
}

function transformPastedHTML(html) {
    if (html.indexOf(`microsoft-com`) !== -1 && html.indexOf(`office`) !== -1) {
        console.log(`transforming`);
        html = transformLists(html);
        html = transformRemoveBookmarks(html);
        html = transformMsoStyles(html);
    }
    return html;
}

function transformMsoStyles(html) {
    html = html.replace(/<o:p>(.*?)<\/o:p>/g, ``);

    const parser = new DOMParser();
    const doc = parser.parseFromString(html, `text/html`);
    doc.querySelectorAll(`[style*="mso-"]`).forEach(node => {
        const styles = parseStyleAttribute(node);
        const newStyles = [];
        for (const prop of Object.keys(styles)) {
            if (prop && !prop.startsWith(`mso-`)) {
                newStyles.push(`${prop}: ${styles[prop]}`);
            }
        }
        node.setAttribute(`style`, newStyles.join(`;`));
    });

    doc.querySelectorAll(`[style*="color: black"]`).forEach(node => {
        node.style.removeProperty(`color`);
    });

    return doc.documentElement.outerHTML;
}

function transformRemoveBookmarks(html) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, `text/html`);
    const bookmarks = doc.querySelectorAll(`[style*="mso-bookmark:"]`);
    bookmarks.forEach(node => {
        const bookmark = parseStyleAttribute(node)[`mso-bookmark`];
        const bookmarkLink = doc.querySelector(`a[name="${bookmark}"]`);
        if (bookmarkLink) {
            bookmarkLink.parentNode.removeChild(bookmarkLink);
        }
        unwrapNode(node);
    });

    return doc.documentElement.outerHTML;
}

function transformLists(html) {
    if (html.indexOf(`mso-list:`) === -1) {
        return html;
    }

    const parser = new DOMParser();
    const doc = parser.parseFromString(html, `text/html`);

    let listStack = [];
    let currentListId;
    const listElements = doc.querySelectorAll(`p[style*="mso-list:"]`);
    listElements.forEach(node => {
        const el = node;
        const [msoListId, msoListLevel] = parseMsoListAttribute(parseStyleAttribute(el)[`mso-list`]);

        // Check for start of a new list
        if (currentListId !== msoListId && (hasNonListItemSibling(el) || msoListLevel === 1)) {
            currentListId = msoListId;
            listStack = [];
        }

        while (msoListLevel > listStack.length) {
            const newList = createListElement(el);

            if (listStack.length > 0) {
                listStack[listStack.length - 1].appendChild(newList);
            } else {
                el.before(newList);
            }
            listStack.push(newList);
        }

        while (msoListLevel < listStack.length) {
            listStack.pop();
        }

        // Remove list item numbers and create li
        listStack[listStack.length - 1].appendChild(getListItemFromParagraph(el));
        el.remove();
    });

    return doc.documentElement.outerHTML;
}

function hasNonListItemSibling(el) {
    return (
        !el.previousElementSibling ||
        !(el.previousElementSibling.nodeName === `OL` || el.previousElementSibling.nodeName === `UL`)
    );
}

function getListItemFromParagraph(el) {
    const li = document.createElement(`li`);
    li.innerHTML = el.innerHTML.replace(listTypeRegex, ``);

    return li;
}

// Parses `mso-list` style attribute
function parseMsoListAttribute(attr) {
    const msoListValue = attr;
    const msoListInfos = msoListValue.split(` `);
    const msoListId = msoListInfos.find(e => /l[0-9]+/.test(e));
    const msoListLevel = +msoListInfos.find(e => e.startsWith(`level`))?.substring(5) || 1;

    return [msoListId, msoListLevel];
}

const listTypeRegex = /<!--\[if \!supportLists\]-->((.|\n)*)<!--\[endif\]-->/m;
function getListPrefix(el) {
    const matches = el.innerHTML.match(listTypeRegex);
    if (matches?.length) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(matches[0], `text/html`);
        return doc.body.querySelector(`span`).textContent;
    }

    return ``;
}

function parseStyleAttribute(el) {
    const styleRaw = el?.attributes[`style`]?.value || ``;
    return Object.fromEntries(styleRaw.split(`;`).map(line => line.split(`:`).map(v => v.trim())));
}

function createListElement(el) {
    const listInfo = getListInfo(getListPrefix(el));
    const list = document.createElement(listInfo.type);
    if (listInfo.countType) {
        list.setAttribute(`type`, listInfo.countType);
    }
    if (listInfo.start > 1) {
        list.setAttribute(`start`, listInfo.start.toString());
    }
    return list;
}

const listOrderRegex = {
    number: /[0-9]+\./,
    romanLower: /(?=[mdclxvi])m*(c[md]|d?c*)(x[cl]|l?x*)(i[xv]|v?i*)\./,
    romanUpper: /(?=[MDCLXVI])M*(C[MD]|D?C*)(X[CL]|L?X*)(I[XV]|V?I*)\./,
    letterLower: /[a-z]+\./,
    letterUpper: /[A-Z]+\./
};

function getListInfo(prefix) {
    let type = `ul`;
    let countType = null;
    let start = 1;
    if (listOrderRegex.number.test(prefix)) {
        type = `ol`;
        start = +prefix.match(listOrderRegex.number)[0].replace(`.`, ``);
    } else if (listOrderRegex.romanLower.test(prefix)) {
        type = `ol`;
        countType = `i`;
        start = +parseRomanNumber(prefix.match(listOrderRegex.romanLower)[0].replace(`.`, ``));
    } else if (listOrderRegex.romanUpper.test(prefix)) {
        type = `ol`;
        countType = `I`;
        start = +parseRomanNumber(prefix.match(listOrderRegex.romanUpper)[0].replace(`.`, ``));
    } else if (listOrderRegex.letterLower.test(prefix)) {
        type = `ol`;
        countType = `a`;
        start = +parseLetterNumber(prefix.match(listOrderRegex.letterLower)[0].replace(`.`, ``));
    } else if (listOrderRegex.letterUpper.test(prefix)) {
        type = `ol`;
        countType = `A`;
        start = +parseLetterNumber(prefix.match(listOrderRegex.letterUpper)[0].replace(`.`, ``));
    }

    return {
        type,
        start,
        countType
    };
}
