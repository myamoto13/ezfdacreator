package com.extia.io.excel.data;

public class Cable {
	
	private String equipement;
	private String slot;
	private String port;
	private String positionAuNRO;
	private String cableDeDistrib;
	private String couleurTube;
	private String fibre;
	private String couleurFibre;
	private String splitter;
	private String tray;
	private String couleurFibre2;
	private String fibre2;
	private String couleurTube2;
	private String cableRaccordement;
	
	public String getEquipement() {
		return equipement;
	}
	public void setEquipement(String equipement) {
		this.equipement = equipement;
	}
	public String getSlot() {
		return slot;
	}
	public void setSlot(String slot) {
		this.slot = slot;
	}
	public String getPort() {
		return port;
	}
	public void setPort(String port) {
		this.port = port;
	}
	public String getPositionAuNRO() {
		return positionAuNRO;
	}
	public void setPositionAuNRO(String positionAuNRO) {
		this.positionAuNRO = positionAuNRO;
	}
	public String getCableDeDistrib() {
		return cableDeDistrib;
	}
	public void setCableDeDistrib(String cableDeDistrib) {
		this.cableDeDistrib = cableDeDistrib;
	}
	public String getCouleurTube() {
		return couleurTube;
	}
	public void setCouleurTube(String couleurTube) {
		this.couleurTube = couleurTube;
	}
	public String getFibre() {
		return fibre;
	}
	public void setFibre(String fibre) {
		this.fibre = fibre;
	}
	public String getCouleurFibre() {
		return couleurFibre;
	}
	public void setCouleurFibre(String couleurFibre) {
		this.couleurFibre = couleurFibre;
	}
	public String getSplitter() {
		return splitter;
	}
	public void setSplitter(String splitter) {
		this.splitter = splitter;
	}
	public String getTray() {
		return tray;
	}
	public void setTray(String tray) {
		this.tray = tray;
	}
	
	public String getCouleurFibre2() {
		return couleurFibre2;
	}
	public void setCouleurFibre2(String couleurFibre2) {
		this.couleurFibre2 = couleurFibre2;
	}
	public String getFibre2() {
		return fibre2;
	}
	public void setFibre2(String fibre2) {
		this.fibre2 = fibre2;
	}
	public String getCouleurTube2() {
		return couleurTube2;
	}
	public void setCouleurTube2(String couleurTube2) {
		this.couleurTube2 = couleurTube2;
	}
	public String getCableRaccordement() {
		return cableRaccordement;
	}
	public void setCableRaccordement(String cableRaccordement) {
		this.cableRaccordement = cableRaccordement;
	}
	
	@Override
	public String toString() {
		return "Cable [equipement=" + equipement + ", slot=" + slot + ", port="
				+ port + ", positionAuNRO=" + positionAuNRO
				+ ", cableDeDistrib=" + cableDeDistrib + ", couleurTube="
				+ couleurTube + ", fibre=" + fibre + ", couleurFibre="
				+ couleurFibre + ", splitter=" + splitter + ", tray=" + tray
				+ ", couleurFibre2=" + couleurFibre2 + ", fibre2=" + fibre2
				+ ", couleurTube2=" + couleurTube2 + ", cableRaccordement="
				+ cableRaccordement + "]";
	}
	
}
