import * as React from "react";
import PropTypes from "prop-types";
import { Image, Text, makeStyles, tokens, shorthands, Button } from "@fluentui/react-components";
import { DocumentBulletListRegular, DeleteDismissRegular, ArrowSyncRegular } from "@fluentui/react-icons";

const useStyles = makeStyles({
  header: {
    display: "flex",
    alignItems: "center",
    ...shorthands.gap("12px"),
    ...shorthands.padding("16px"),
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundInverted,
    justifyContent: "space-between",
  },
  title: {
    margin: 0,
    fontSize: "18px",
    fontWeight: "600",
    color: tokens.colorNeutralForegroundInverted,
  },
  subtitle: {
    color: tokens.colorNeutralForegroundInvertedLink,
    fontSize: "12px",
  },
  logo: {
    width: "28px",
    height: "28px",
    ...shorthands.borderRadius("4px"),
    ...shorthands.padding("2px"),
    backgroundColor: tokens.colorNeutralForegroundInverted,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
  },
  logoImage: {
    width: "20px",
    height: "20px",
  },
  headerLeft: {
    display: "flex",
    alignItems: "center",
    ...shorthands.gap("12px"),
  },
  headerRight: {
    display: "flex",
    alignItems: "center",
    ...shorthands.gap("8px"),
  },
  headerButton: {
    color: tokens.colorNeutralForegroundInverted,
    "&:hover": {
      backgroundColor: "rgba(255, 255, 255, 0.1)",
    }
  },
  timer: {
    color: tokens.colorNeutralForegroundInvertedLink,
    fontSize: "12px",
    marginTop: "4px",
  },
});

const Header = (props) => {
  const { title, logo, message, onClear, onAnalyse, isProcessing, analysisTime } = props;
  const styles = useStyles();

  return (
    <div className={styles.header}>
      <div className={styles.headerLeft}>
        <div className={styles.logo}>
          {logo ? (
            <Image className={styles.logoImage} src={logo} alt={title} />
          ) : (
            <DocumentBulletListRegular />
          )}
        </div>
        <div>
          <h1 className={styles.title}>{message}</h1>
          <Text size={200} className={styles.subtitle}>{title}</Text>
          {analysisTime > 0 && (
            <Text size={200} className={styles.timer}>
              Analysis completed in {analysisTime.toFixed(2)} seconds
            </Text>
          )}
        </div>
      </div>
      <div className={styles.headerRight}>
        {onAnalyse && (
          <Button
            appearance="transparent"
            className={styles.headerButton}
            onClick={onAnalyse}
            disabled={isProcessing}
            icon={<ArrowSyncRegular />}
          >
            Analyse
          </Button>
        )}
        {onClear && (
          <Button
            appearance="transparent"
            className={styles.headerButton}
            onClick={onClear}
            disabled={isProcessing}
            icon={<DeleteDismissRegular />}
          >
            Clear
          </Button>
        )}
      </div>
    </div>
  );
};

Header.propTypes = {
  title: PropTypes.string,
  logo: PropTypes.string,
  message: PropTypes.string,
  onClear: PropTypes.func,
  onAnalyse: PropTypes.func,
  isProcessing: PropTypes.bool,
  analysisTime: PropTypes.number,
};

export default Header;
