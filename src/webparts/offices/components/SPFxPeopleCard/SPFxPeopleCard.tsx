import * as React from "react";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { ServiceScope } from "@microsoft/sp-core-library";
import {
    Persona,
    PersonaSize
} from "@fluentui/react";

import styles from "./SPFxPeopleCard.module.scss";

export interface IPeopleCardProps {
    primaryText: string;
    secondaryText?: string;
    tertiaryText?: string;
    optionalText?: string;
    moreDetail?: HTMLElement | string;
    pictureUrl?: string;
    email: string;
    serviceScope: ServiceScope;
    class?: string;
    size: PersonaSize;
    width?: number;
    height?: number;
}

const LIVE_PERSONA_COMPONENT_ID: string = "914330ee-2df2-4f6e-a858-30c23a812408";

const headshotUrls = {
    delve: `https://nam.delve.office.com/mt/v3/people/profileimage?userId={{EMAIL}}&size=L`,
    spo: `/_layouts/15/userphoto.aspx?size=L&accountname={{EMAIL}}`,
    spoSearch: ``,
    graph: ``,
    outlook: `https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email={{EMAIL}}&UA=0&size=HR96x96`,
};

export const getHeadshotUrl = (type: keyof typeof headshotUrls, email: string, spoHeadshotUrl?: string): string => {
    if (type === "spoSearch" || type === "graph") {
        return spoHeadshotUrl || headshotUrls.spo.replace("{{EMAIL}}", encodeURIComponent(email));
    }

    return (headshotUrls[type] || headshotUrls.spo).replace("{{EMAIL}}", encodeURIComponent(email));
};

const SPFxPeopleCard: React.FunctionComponent<IPeopleCardProps> = (props: IPeopleCardProps) => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const [personaCard, setPersonaCard] = React.useState<any>({});
    const [imgSrc, setImgSrc] = React.useState<string>(props.pictureUrl || getHeadshotUrl("spo", props.email));

    const _imageError = (): void => {
        setImgSrc(getHeadshotUrl("spo", props.email));
    };

    /**
     * Display default OfficeUIFabric Persona card if SPFx LivePersonaCard not loaded
     */
    const _defaultContactCard = (): JSX.Element => {
        if (props.primaryText) {
            return (
                <Persona
                    size={props.size}
                    imageUrl={imgSrc}
                    onError={_imageError}
                    text={props.primaryText}
                    secondaryText={props.secondaryText}
                />
            );
        } else {
            return (
                <div
                    className={styles.headshotWrap}
                    style={{
                        width: props.width || 72,
                        height: props.height || 72,
                        backgroundImage: `url('${imgSrc}')`,
                    }}
                >
                    <img src={imgSrc} onError={_imageError} />
                </div>
            );
        }
    };

    /**
     * Configure SPFx LivePersona card from SPFx component loader
     */
    const _spfxLiverPersonaCard = (): JSX.Element => {
        return React.createElement(
            personaCard.card,
            {
                className: "people",
                clientScenario: "PeopleWebPart",
                disableHover: false,
                hostAppPersonaInfo: {
                    PersonaType: "User",
                },
                serviceScope: props.serviceScope,
                upn: props.email,
                legacyUpn: props.email,
            },
            _defaultContactCard()
        );
    };

    React.useEffect(() => {
        SPComponentLoader.loadComponentById(LIVE_PERSONA_COMPONENT_ID)
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            .then((sharedLibrary: any) => {
                setPersonaCard({ card: sharedLibrary.LivePersonaCard });
            })
            .catch((er) => console.log("Error loading SP Component", er));
    }, []);

    return (
        <div className={`${props.class || ``} ${!props.primaryText ? styles.imageOnly : ``}`}>
            {personaCard.card ? _spfxLiverPersonaCard() : _defaultContactCard()}
        </div>
    );
};

export default SPFxPeopleCard;
