import * as React from "react";
import PropTypes from "prop-types";
import { Image, Text, makeStyles } from "@fluentui/react-components";

const useStyles = makeStyles({
  header: {
    display: "flex",
    alignItems: "center",
    gap: "10px",
    paddingBottom: "15px",
    marginBottom: "15px",
    borderBottom: "1px solid #e0e0e0",
  },
  title: {
    margin: 0,
    fontSize: "20px",
    fontWeight: "600",
  },
  logo: {
    width: "32px",
    height: "32px",
  }
});

const Header = (props) => {
  const { title, logo, message } = props;
  const styles = useStyles();

  return (
    <div className={styles.header}>
      <Image className={styles.logo} src={logo} alt={title} />
      <div>
        <h1 className={styles.title}>{message}</h1>
        <Text size={200}>{title}</Text>
      </div>
    </div>
  );
};

Header.propTypes = {
  title: PropTypes.string,
  logo: PropTypes.string,
  message: PropTypes.string,
};

export default Header;
