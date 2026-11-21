// Games and raffle content edited in home.exe's Wedding games app, kept in Firestore's weddingGames collection.
// Anyone may read it without signing in (the host's phone, the LED laptop); only the household can write.
// A Firebase web key is public by design; the Firestore rules decide who may write.
const WEDDING_GAMES_URL = 'https://firestore.googleapis.com/v1/projects/home-exe/databases/(default)/documents/weddingGames?pageSize=300&key=AIzaSyAmKq2v_kEWJpWQ7riKIAhX93fYQIYiQMg';

// Resolves to every list, each in its saved order; rejects when Firestore can't be reached or refuses,
// so each page can fall back to its JSON file.
function weddingGamesLoad() {
  const value = v => 'stringValue' in v ? v.stringValue : 'integerValue' in v ? Number(v.integerValue) : 'doubleValue' in v ? v.doubleValue : 'booleanValue' in v ? v.booleanValue : '';
  return fetch(WEDDING_GAMES_URL, { cache: 'no-store' }).then(r => { if (!r.ok) throw new Error('Firestore ' + r.status); return r.json(); }).then(res => {
    const items = (res.documents || []).map(doc => Object.entries(doc.fields || {}).reduce((o, [k, v]) => ({ ...o, [k]: value(v) }), { id: doc.name.split('/').pop() }));
    const pick = prefix => items.filter(i => i.id.startsWith(prefix)).sort((a, b) => (a.order || 0) - (b.order || 0));
    return { questions: pick('q-'), tiebreakers: pick('tb-'), songs: pick('song-'), prizes: pick('prize-') };
  });
}
