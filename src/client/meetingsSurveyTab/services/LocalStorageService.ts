import { dateAdd, PnPClientStorage, PnPClientStore } from "@pnp/common";

export class LocalStorageService {

    static Common = class Common {

        static async putAsync(localStorage: PnPClientStore, key: string, getter: () => Promise<any>, expire: Date): Promise<void> {
            localStorage.put(key, await getter(), expire)
        }

        static put(localStorage: PnPClientStore, key: string, obj: any, expire: Date): void {
            localStorage.put(key, obj, expire)
        }
    };

    static ExpirationConstants = class ExpirationConstants {
        static Days(days: number): Date {
            return dateAdd(new Date(), 'day', days)
        }

        static Hours(hours: number): Date {
            return dateAdd(new Date(), 'hour', hours)
        }

        static Minutes(minutes: number): Date {
            return dateAdd(new Date(), 'minute', minutes)
        }

        static Seconds(seconds: number): Date {
            return dateAdd(new Date(), 'second', seconds)
        }
    };

    static AuthToken = class AuthToken {
        static GetAuthTokenKey = "meetings-survey-auth-token";

        static deleteAll(): void {
            LocalStorageService.deleteLocalStorageCache([
                AuthToken.GetAuthTokenKey,
            ])
        }

        static async getAuthToken(userId: string, getter: () => Promise<string>, needBackgroundUpdateCache: boolean = false, minutes: number = 30): Promise<string> {
            const storage = new PnPClientStorage();
            await storage.local.deleteExpired();
            let expire = LocalStorageService.ExpirationConstants.Minutes(minutes);
            let result = await storage.local.getOrPut(`${AuthToken.GetAuthTokenKey}-${userId}`, getter, expire);
            if (needBackgroundUpdateCache)
                LocalStorageService.Common.putAsync(
                    storage.local,
                    `${AuthToken.GetAuthTokenKey}-${userId}`,
                    getter,
                    expire,
                );
            return result;
        }
    };

    static Config = class Config {
        static GetConfigKey = "meetings-survey-configuration";

        static deleteAll(): void {
            LocalStorageService.deleteLocalStorageCache([
                Config.GetConfigKey,
            ])
        }

        static async getConfig(getter: () => Promise<object>, needBackgroundUpdateCache: boolean = false, hours: number = 1): Promise<object> {
            const storage = new PnPClientStorage();
            await storage.local.deleteExpired();
            let expire = LocalStorageService.ExpirationConstants.Hours(hours);
            let result = await storage.local.getOrPut(Config.GetConfigKey, getter, expire);
            if (needBackgroundUpdateCache)
                LocalStorageService.Common.putAsync(
                    storage.local,
                    Config.GetConfigKey,
                    getter,
                    expire,
                );
            return result;
        }
    };

    static deleteLocalStorageCache(keys: string[]) {
        const storage = new PnPClientStorage();
        if (!!keys && keys.length > 0) {
            const localStorageKeys = Object.keys(window.localStorage);
            keys.forEach(key => {
                localStorageKeys.forEach(lsKey => {
                    if (!!lsKey && lsKey.indexOf(key) > -1)
                        storage.local.delete(lsKey)
                })
            })
        }
    }
}
