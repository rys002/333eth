pragma solidity ^0.4.23;

import "../Revolution.sol";


contract TestRejectingInvestor {
  bool public acceptsPayments;

  function() external payable {
    require(acceptsPayments, "payment rejected");
  }

  function invest(Revolution revolution) external payable {
    address[3] memory refs;
    revolution.doInvest.value(msg.value)(refs);
  }

  function claim(Revolution revolution) external {
    revolution.getMyDividends();
  }

  function setAcceptsPayments(bool accepts) external {
    acceptsPayments = accepts;
  }
}
