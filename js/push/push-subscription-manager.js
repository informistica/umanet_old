// urlB64ToUint8Array is a magic function that will encode the base64 public key
// to Array buffer which is needed by the subscription option
const urlB64ToUint8Array = base64String => {
    const padding = '='.repeat((4 - (base64String.length % 4)) % 4)
    const base64 = (base64String + padding).replace(/\-/g, '+').replace(/_/g, '/')
    const rawData = atob(base64)
    const outputArray = new Uint8Array(rawData.length)
    for (let i = 0; i < rawData.length; ++i) {
        outputArray[i] = rawData.charCodeAt(i)
    }
    return outputArray
}

function buildPostBody(obj) {
    const formData = new FormData();
    for (const key in obj) {
        formData.append(key, obj[key]);
    }

    return new URLSearchParams(formData);
}

const getCodiceAllievo = () => {
    return document.getElementById('CodiceAllievo').value;
};

// controlla che tutte le feature necessarie siano supportate
const check = () => {
    if (!('serviceWorker' in navigator)) {
        // throw new Error('No Service Worker support!')
        return false;
    }
    if (!('PushManager' in window)) {
        // throw new Error('No Push API Support!')
        return false;
    }
    if (!("Notification" in window)) {
        // throw new Error('No Notification API Support!')
        return false;
    }

    return true;
}

// chiede all'utente il permesso per le notifiche
const requestNotificationPermission = async () => {
    const permission = await window.Notification.requestPermission()
    // value of permission can be 'granted', 'default', 'denied'
    // granted: user has accepted the request
    // default: user has dismissed the notification permission popup by clicking on x
    // denied: user has denied the request.
    if (permission !== 'granted') {
        throw new Error('Permission not granted for Notification')
    }
}

// crea iscrizione
const makeSubscription = async () => {
    /*
    const applicationServerKey = urlB64ToUint8Array(
        // NOTA: deve essere la stessa chaive pubblica che è in .env
        'BERKNfKKdoJ47h4bkQ3vsu40hyEy4yJkv_YlHGOx2IHPMSZ87V0gcCqoijnq9CYl7KqOKE17ZnTBIzEguTLlfMk'
    );
    */
    const options = {
        applicationServerKey: 'BERKNfKKdoJ47h4bkQ3vsu40hyEy4yJkv_YlHGOx2IHPMSZ87V0gcCqoijnq9CYl7KqOKE17ZnTBIzEguTLlfMk',
        userVisibleOnly: true,
    };

    const registration = await navigator.serviceWorker.ready;
    return registration.pushManager.subscribe(options);
}

const saveSubscription = subscription => {
    const { endpoint, expirationTime, keys } = subscription.toJSON();
    const { auth, p256dh } = keys;

    const body = {
        CodiceAllievo: getCodiceAllievo(),
        endpoint,
        auth,
        p256dh,
    };
    if (expirationTime)
        body.expirationTime = expirationTime;

    return fetch('/googleapi/save-subscription.php', {
        method: 'post',
        body: buildPostBody(body),
    });
}

// ottiene la registrazione o si iscrive
const subscribe = () => navigator.serviceWorker.ready
    .then(registration => registration.pushManager.getSubscription())
    .then(subscription => subscription ? subscription : makeSubscription())
    .then(subscription => saveSubscription(subscription))

// TODO: mostrare erroe se il service worker non è registrato o se browser non supporta feature
const isSubscribed = async () => {
    const registration = await navigator.serviceWorker.ready;
    const subscription = await registration.pushManager.getSubscription();

    if (!subscription)
        return false;

    const res = await fetch('/googleapi/is-subscribed.php', {
        method: 'post',
        body: buildPostBody({
            //CodiceAllievo: 'darcros',
            CodiceAllievo:getCodiceAllievo(),
            endpoint: subscription.endpoint,
        }),
    });
    const { isSubscribed: serverSubscribed } = await res.json();
    return !!serverSubscribed;
}

async function onLoad() {
    const subscriptionStatusField = document.getElementById('push-subscription-status');
    const btn = document.getElementById('push-subscribe-button');
    if (!check()) {
        subscriptionStatusField.value = "Notifiche push non supportate da questo browser";
        btn.disabled = true;
        return;
    }

    const iscritto = await isSubscribed();

    subscriptionStatusField.value = `${iscritto ? "" : "non "} iscritto su questo dispositivo`;
    btn.disabled = iscritto;
}
document.addEventListener('DOMContentLoaded', () => onLoad());

async function subscribeButton() {
    const serverRes = await subscribe();
    if (serverRes.ok) {
        alert('iscrizione salvata');
        location.reload();
    } else {
        alert('errore durante il savaggio dell\'iscrizione. Riprova.');
    }
}
